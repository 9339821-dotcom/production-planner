
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os
import sys
from collections import defaultdict
import json
import re

class ProductionPlanner:
    def __init__(self, excel_file):
        self.excel_file = excel_file
        self.orders_df = None
        self.materials_df = None
        self.stock_data = {}
        self.reserved_materials = defaultdict(float)
        self.selected_orders = {}
        self.load_data()
    
    def load_data(self):
        """Загрузка данных из Excel файла"""
        try:
            print("📂 Загрузка данных из Excel файла...")
            
            # Загружаем лист с заказами
            self.orders_df = pd.read_excel(self.excel_file, sheet_name='Заказы')
            
            # Загружаем лист с материалами
            self.materials_df = pd.read_excel(self.excel_file, sheet_name='Потребность материалов')
            
            # Создаем словарь остатков на складе из колонки "На складе"
            if 'На складе' in self.materials_df.columns:
                for _, row in self.materials_df.iterrows():
                    material = row['Материал']
                    if pd.notna(material):
                        stock = row['На складе'] if pd.notna(row['На складе']) else 0
                        self.stock_data[str(material).strip()] = float(stock)
            
            print(f"✅ Загружено: {len(self.orders_df)} заказов, {len(self.materials_df)} материалов")
            
        except Exception as e:
            print(f"❌ Ошибка загрузки данных: {e}")
            raise
    
    def get_companies(self):
        """Получить список компаний"""
        return sorted([str(x) for x in self.orders_df['Клиент'].unique() if pd.notna(x)])
    
    def get_orders_by_company(self, company=None):
        """Получить заказы по компании"""
        if company and company != "Все компании":
            return self.orders_df[self.orders_df['Клиент'] == company]
        return self.orders_df
    
    def select_orders(self, order_numbers, shipment_dates):
        """Выбрать заказы для планирования"""
        self.selected_orders = {}
        
        for order_num in order_numbers:
            # Преобразуем номер заказа к строке для сравнения
            order_num_str = str(order_num).strip()
            order_data = self.orders_df[self.orders_df['Номер заказа'].astype(str).str.strip() == order_num_str]
            
            if not order_data.empty:
                order_info = order_data.iloc[0].to_dict()
                order_info['Дата отгрузки'] = shipment_dates.get(order_num_str)
                self.selected_orders[order_num_str] = order_info
        
        print(f"✅ Выбрано заказов: {len(self.selected_orders)}")
        return self.selected_orders
    
    def calculate_material_requirements(self):
        """Рассчитать потребность в материалах для выбранных заказов"""
        if not self.selected_orders:
            return {"error": "Не выбраны заказы"}
        
        required_materials = defaultdict(float)
        order_materials = {}
        
        # Получаем все материалы из таблицы потребности
        all_materials = []
        if 'Материал' in self.materials_df.columns:
            all_materials = [str(x).strip() for x in self.materials_df['Материал'] if pd.notna(x)]
        
        # Для каждого выбранного заказа находим его материалы
        for order_num in self.selected_orders.keys():
            order_num_clean = str(order_num).strip()
            
            # Ищем колонку с этим заказом в таблице материалов
            order_columns = []
            for col in self.materials_df.columns:
                col_str = str(col).strip()
                # Ищем точное совпадение или частичное (если есть дополнительные символы)
                if order_num_clean == col_str or order_num_clean in col_str.split():
                    order_columns.append(col)
            
            if order_columns:
                for material_idx, material_name in enumerate(all_materials):
                    total_requirement = 0
                    
                    for order_col in order_columns:
                        if order_col in self.materials_df.columns:
                            value = self.materials_df[order_col].iloc[material_idx]
                            if pd.notna(value) and value != 0:
                                try:
                                    total_requirement += float(value)
                                except (ValueError, TypeError):
                                    pass
                    
                    if total_requirement > 0:
                        required_materials[material_name] += total_requirement
                        if order_num not in order_materials:
                            order_materials[order_num] = {}
                        order_materials[order_num][material_name] = total_requirement
        
        # Рассчитываем остатки после резервирования
        material_balance = {}
        purchase_requirements = {}
        
        for material, required in required_materials.items():
            current_stock = self.stock_data.get(material, 0)
            reserved = self.reserved_materials.get(material, 0)
            available_stock = max(0, current_stock - reserved)
            
            balance_after = available_stock - required
            material_balance[material] = {
                'Текущий запас': current_stock,
                'Уже зарезервировано': reserved,
                'Доступно сейчас': available_stock,
                'Требуется для выбранных': required,
                'Остаток после': balance_after
            }
            
            # Если будет дефицит - добавляем в заявку на закупку
            if balance_after < 0:
                purchase_requirements[material] = abs(balance_after)
        
        return {
            'material_requirements': dict(required_materials),
            'material_balance': material_balance,
            'purchase_requirements': purchase_requirements,
            'order_materials': order_materials
        }
    
    def reserve_materials(self, order_numbers, shipment_dates):
        """Зарезервировать материалы для заказов"""
        # Выбираем заказы
        self.select_orders(order_numbers, shipment_dates)
        
        # Рассчитываем потребности
        requirements = self.calculate_material_requirements()
        
        if 'error' in requirements:
            return requirements
        
        # Резервируем материалы
        for material, required in requirements['material_requirements'].items():
            self.reserved_materials[material] += required
        
        # Сохраняем информацию о резервировании
        self.save_reservation_data()
        
        return {
            'status': 'success',
            'reserved_orders': list(self.selected_orders.keys()),
            'reserved_materials': dict(self.reserved_materials),
            'requirements': requirements
        }
    
    def save_reservation_data(self):
        """Сохранить данные о резервировании"""
        # Конвертируем даты в строки для JSON
        serializable_orders = {}
        for order_num, order_info in self.selected_orders.items():
            serializable_order = order_info.copy()
            if 'Дата отгрузки' in serializable_order and serializable_order['Дата отгрузки']:
                if isinstance(serializable_order['Дата отгрузки'], datetime):
                    serializable_order['Дата отгрузки'] = serializable_order['Дата отгрузки'].isoformat()
            serializable_orders[order_num] = serializable_order
        
        reservation_data = {
            'reserved_materials': dict(self.reserved_materials),
            'selected_orders': serializable_orders,
            'timestamp': datetime.now().isoformat()
        }
        
        try:
            with open('reservations.json', 'w', encoding='utf-8') as f:
                json.dump(reservation_data, f, ensure_ascii=False, indent=2)
            print("💾 Данные резервирования сохранены")
        except Exception as e:
            print(f"⚠️ Не удалось сохранить данные резервирования: {e}")
    
    def load_reservation_data(self):
        """Загрузить данные о резервировании"""
        try:
            if os.path.exists('reservations.json'):
                with open('reservations.json', 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.reserved_materials = defaultdict(float, data.get('reserved_materials', {}))
                    
                    # Восстанавливаем даты из строк
                    loaded_orders = data.get('selected_orders', {})
                    for order_num, order_info in loaded_orders.items():
                        if 'Дата отгрузки' in order_info and order_info['Дата отгрузки']:
                            try:
                                order_info['Дата отгрузки'] = datetime.fromisoformat(order_info['Дата отгрузки'])
                            except (ValueError, TypeError):
                                order_info['Дата отгрузки'] = None
                        self.selected_orders[order_num] = order_info
                    
                    print(f"📂 Загружены предыдущие резервирования: {len(self.selected_orders)} заказов")
                    return True
        except Exception as e:
            print(f"⚠️ Не удалось загрузить данные резервирования: {e}")
        return False
    
    def clear_reservations(self):
        """Очистить все резервирования"""
        self.reserved_materials.clear()
        self.selected_orders.clear()
        try:
            if os.path.exists('reservations.json'):
                os.remove('reservations.json')
                print("🗑️ Файл резервирования удален")
        except Exception as e:
            print(f"⚠️ Не удалось удалить файл резервирования: {e}")
        return {"status": "success", "message": "Все резервирования очищены"}
    
    def generate_purchase_order(self, requirements):
        """Сформировать заявку на закупку"""
        if not requirements.get('purchase_requirements'):
            return "Заявка на закупку не требуется - все материалы в наличии", None
        
        purchase_text = "ЗАЯВКА НА ЗАКУПКУ МАТЕРИАЛОВ\n"
        purchase_text += "=" * 50 + "\n"
        purchase_text += f"Дата формирования: {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
        purchase_text += f"Для заказов: {', '.join(self.selected_orders.keys())}\n\n"
        
        purchase_text += "СПИСОК МАТЕРИАЛОВ ДЛЯ ЗАКУПКИ:\n"
        purchase_text += "-" * 50 + "\n"
        
        total_cost_estimate = 0
        
        for material, quantity in requirements['purchase_requirements'].items():
            estimated_price = self.estimate_material_price(material)
            total_cost = estimated_price * quantity
            total_cost_estimate += total_cost
            
            purchase_text += f"📦 {material}\n"
            purchase_text += f"   Количество: {quantity:.2f}\n"
            purchase_text += f"   Примерная стоимость: {total_cost:,.2f} руб.\n\n"
        
        purchase_text += f"ОБЩАЯ ПРИМЕРНАЯ СТОИМОСТЬ: {total_cost_estimate:,.2f} руб.\n"
        
        # Сохраняем в файл
        filename = f"Заявка_на_закупку_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
        try:
            with open(filename, 'w', encoding='utf-8') as f:
                f.write(purchase_text)
            return purchase_text, filename
        except Exception as e:
            print(f"❌ Ошибка сохранения заявки: {e}")
            return purchase_text, None
    
    def estimate_material_price(self, material):
        """Оценочная стоимость материала"""
        material_lower = material.lower()
        
        # Базовые цены на основные материалы
        if any(x in material_lower for x in ['стекло', 'glass']):
            return 1500
        elif any(x in material_lower for x in ['профиль', 'profile']):
            return 800
        elif any(x in material_lower for x in ['аргон', 'argon']):
            return 200
        elif any(x in material_lower for x in ['герметик', 'sealant']):
            return 1500
        elif any(x in material_lower for x in ['лента', 'tape']):
            return 300
        elif any(x in material_lower for x in ['соединитель', 'connector']):
            return 500
        else:
            return 1000  # цена по умолчанию

def main():
    """Основная функция для консольного интерфейса"""
    print("🏭 СИСТЕМА ПЛАНИРОВАНИЯ ПРОИЗВОДСТВА")
    print("=" * 50)
    
    # Проверяем наличие файла
    excel_file = "Объединенная_статистика_заказов.xlsx"
    if not os.path.exists(excel_file):
        print(f"❌ Файл {excel_file} не найден!")
        print("Поместите файл в ту же папку, что и программу.")
        input("Нажмите Enter для выхода...")
        return
    
    # Создаем планировщик
    try:
        planner = ProductionPlanner(excel_file)
    except Exception as e:
        print(f"❌ Не удалось загрузить данные: {e}")
        input("Нажмите Enter для выхода...")
        return
    
    # Загружаем предыдущие резервирования
    planner.load_reservation_data()
    
    while True:
        print("\n" + "=" * 50)
        print("📋 ГЛАВНОЕ МЕНЮ")
        print("1. Показать все заказы")
        print("2. Выбрать заказы для планирования")
        print("3. Показать текущие резервирования")
        print("4. Рассчитать потребности в материалах")
        print("5. Сформировать заявку на закупку")
        print("6. Очистить все резервирования")
        print("7. Выход")
        
        choice = input("\nВыберите действие (1-7): ").strip()
        
        if choice == '1':
            show_orders_menu(planner)
        elif choice == '2':
            select_orders_menu(planner)
        elif choice == '3':
            show_reservations(planner)
        elif choice == '4':
            calculate_requirements(planner)
        elif choice == '5':
            generate_purchase_order(planner)
        elif choice == '6':
            result = planner.clear_reservations()
            print(f"✅ {result['message']}")
        elif choice == '7':
            print("👋 Выход из программы...")
            break
        else:
            print("❌ Неверный выбор!")

def show_orders_menu(planner):
    """Меню показа заказов"""
    companies = planner.get_companies()
    print(f"\n🏢 Доступные компании ({len(companies)}):")
    for i, company in enumerate(companies, 1):
        print(f"{i}. {company}")
    print(f"{len(companies) + 1}. Все компании")
    
    try:
        choice = int(input(f"\nВыберите компанию (1-{len(companies) + 1}): "))
        if 1 <= choice <= len(companies):
            selected_company = companies[choice - 1]
        else:
            selected_company = "Все компании"
        
        orders = planner.get_orders_by_company(selected_company)
        print(f"\n📊 Заказы для '{selected_company}' ({len(orders)}):")
        print("-" * 100)
        
        for _, order in orders.iterrows():
            order_num = str(order['Номер заказа'])
            client = str(order['Клиент']) if pd.notna(order['Клиент']) else "Не указан"
            status = str(order.get('Состояние заказа', 'Не указано'))
            cost = order.get('Стоимость заказа', 0)
            area = order.get('Площадь заказа', 0)
            
            print(f"📋 {order_num} | {client} | {status} | {area} м² | {cost:,.2f} руб.")
            
    except (ValueError, IndexError):
        print("❌ Неверный выбор!")

def select_orders_menu(planner):
    """Меню выбора заказов"""
    print("\n🎯 ВЫБОР ЗАКАЗОВ ДЛЯ ПЛАНИРОВАНИЯ")
    
    all_orders = [str(x) for x in planner.orders_df['Номер заказа'].unique() if pd.notna(x)]
    print(f"Всего заказов: {len(all_orders)}")
    
    print("\nПримеры заказов:")
    for order in all_orders[:10]:
        print(f"  - {order}")
    if len(all_orders) > 10:
        print(f"  ... и еще {len(all_orders) - 10}")
    
    order_input = input("\nВведите номера заказов через запятую: ").strip()
    if not order_input:
        print("❌ Не введены номера заказов!")
        return
    
    order_numbers = [num.strip() for num in order_input.split(',') if num.strip()]
    
    # Проверяем существование заказов
    valid_orders = []
    invalid_orders = []
    
    for order_num in order_numbers:
        if any(str(existing_order).strip() == order_num for existing_order in all_orders):
            valid_orders.append(order_num)
        else:
            invalid_orders.append(order_num)
    
    if invalid_orders:
        print(f"❌ Не найдены заказы: {', '.join(invalid_orders)}")
    
    if not valid_orders:
        print("❌ Не выбрано ни одного действительного заказа!")
        return
    
    # Запрашиваем даты отгрузки
    shipment_dates = {}
    print("\n📅 ВВОД ДАТ ОТГРУЗКИ (в формате ДД.ММ.ГГГГ, например: 25.12.2024):")
    print("   Для отмены ввода даты нажмите Enter")
    
    for order_num in valid_orders:
        while True:
            date_str = input(f"Дата отгрузки для заказа {order_num}: ").strip()
            if not date_str:
                use_default = input("  Не указывать дату отгрузки? (y/n): ").strip().lower()
                if use_default == 'y':
                    shipment_dates[order_num] = None
                    break
                else:
                    continue
            
            try:
                shipment_date = datetime.strptime(date_str, '%d.%m.%Y')
                shipment_dates[order_num] = shipment_date
                break
            except ValueError:
                print("❌ Неверный формат даты! Используйте ДД.ММ.ГГГГ")
    
    # Резервируем материалы
    result = planner.reserve_materials(valid_orders, shipment_dates)
    
    if 'status' in result and result['status'] == 'success':
        print(f"\n✅ Успешно зарезервировано материалов для {len(valid_orders)} заказов")
        reserved_count = len(result.get('reserved_materials', {}))
        print(f"📦 Зарезервировано материалов: {reserved_count} позиций")
    else:
        print("❌ Ошибка при резервировании материалов!")

def show_reservations(planner):
    """Показать текущие резервирования"""
    if not planner.reserved_materials:
        print("📭 Нет активных резервирований")
        return
    
    print(f"\n📋 АКТИВНЫЕ РЕЗЕРВИРОВАНИЯ ({len(planner.selected_orders)} заказов):")
    for order_num, order_info in planner.selected_orders.items():
        shipment_date = order_info.get('Дата отгрузки')
        if shipment_date and isinstance(shipment_date, datetime):
            date_str = shipment_date.strftime('%d.%m.%Y')
        else:
            date_str = 'Не указана'
        client = order_info.get('Клиент', 'Не указан')
        print(f"  🚚 {order_num} ({client}) - отгрузка: {date_str}")
    
    print(f"\n📦 ЗАРЕЗЕРВИРОВАННЫЕ МАТЕРИАЛЫ ({len(planner.reserved_materials)} позиций):")
    for material, quantity in list(planner.reserved_materials.items())[:20]:  # Показываем первые 20
        print(f"  📍 {material}: {quantity:.2f}")
    
    if len(planner.reserved_materials) > 20:
        print(f"  ... и еще {len(planner.reserved_materials) - 20} материалов")

def calculate_requirements(planner):
    """Рассчитать потребности в материалах"""
    if not planner.selected_orders:
        print("❌ Сначала выберите заказы!")
        return
    
    print("\n🧮 РАСЧЕТ ПОТРЕБНОСТЕЙ В МАТЕРИАЛАХ")
    
    requirements = planner.calculate_material_requirements()
    
    if 'error' in requirements:
        print(f"❌ {requirements['error']}")
        return
    
    print(f"\n📊 РЕЗУЛЬТАТЫ ДЛЯ {len(planner.selected_orders)} ЗАКАЗОВ:")
    print("=" * 80)
    
    # Показываем баланс по материалам (первые 15)
    materials_shown = 0
    for material, balance in list(requirements['material_balance'].items())[:15]:
        if balance['Требуется для выбранных'] > 0:
            print(f"\n📦 {material}:")
            print(f"   Текущий запас: {balance['Текущий запас']:.2f}")
            print(f"   Уже зарезервировано: {balance['Уже зарезервировано']:.2f}")
            print(f"   Доступно сейчас: {balance['Доступно сейчас']:.2f}")
            print(f"   Требуется для выбранных: {balance['Требуется для выбранных']:.2f}")
            
            remaining = balance['Остаток после']
            if remaining >= 0:
                print(f"   ✅ Остаток после резервирования: {remaining:.2f}")
            else:
                print(f"   ❌ ДЕФИЦИТ: {-remaining:.2f}")
            materials_shown += 1
    
    if len(requirements['material_balance']) > materials_shown:
        print(f"\n... и еще {len(requirements['material_balance']) - materials_shown} материалов")
    
    # Показываем заявку на закупку
    if requirements['purchase_requirements']:
        print(f"\n🚨 ТРЕБУЕТСЯ ЗАКУПКА ({len(requirements['purchase_requirements'])} материалов):")
        for material, quantity in list(requirements['purchase_requirements'].items())[:10]:
            print(f"   📍 {material}: {quantity:.2f}")
        
        if len(requirements['purchase_requirements']) > 10:
            print(f"   ... и еще {len(requirements['purchase_requirements']) - 10} материалов")
    else:
        print(f"\n✅ Все материалы в наличии!")

def generate_purchase_order(planner):
    """Сформировать заявку на закупку"""
    if not planner.selected_orders:
        print("❌ Сначала выберите заказы!")
        return
    
    requirements = planner.calculate_material_requirements()
    
    if 'error' in requirements:
        print(f"❌ {requirements['error']}")
        return
    
    purchase_text, filename = planner.generate_purchase_order(requirements)
    
    print(f"\n✅ ЗАЯВКА НА ЗАКУПКУ СФОРМИРОВАНА:")
    print("=" * 60)
    print(purchase_text)
    print("=" * 60)
    if filename:
        print(f"📄 Сохранено в файл: {filename}")
    else:
        print("📄 Файл не сохранен (ошибка записи)")

if __name__ == "__main__":
    main()
