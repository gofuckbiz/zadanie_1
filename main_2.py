import pandas as pd
import os

def read_excel_file(file_path, required_columns):
    try:
        if not os.path.exists(file_path):
            print(f"[INFO -] Ошибка: Файл '{file_path}' не существует.")
            return None

        data = pd.read_excel(file_path)
        missing_columns = [col for col in required_columns if col not in data.columns]
        if missing_columns:
            print(f"[INFO -] Ошибка: Не существует столбца {missing_columns} в файле '{file_path}'.")
            print(f"Available columns in '{file_path}': {list(data.columns)}")
            return None

        print(f"[INFO +] файл был успешно прочитан '{file_path}'.")
        return data

    except Exception as e:
        print(f"[INFO -] Ошибка при чтении файла '{file_path}': {e}")
        return None

def find_matching_products(supplier_data, product_list):
    try:
        print("[INFO +] Поиск соответствующих продуктов по данным поставщика и списку продуктов...")

        
        matched_products = pd.merge(
            supplier_data, 
            product_list, 
            on='Наименование', 
            how='inner'
        )

        print(f"[INFO +] Успешное сопастовление {len(matched_products)} продуктов из данных поставщика.")
        return matched_products

    except Exception as e:
        print(f"[INFO +] Ошибка при сопоставлении товаров: {e}")
        return None

def validate_with_category_tree(matched_products, category_tree):
    try:
        print("[INFO: PROCESS] Проверка сопоставленных продуктов с помощью дерева категорий...")

       
        validated_products = pd.merge(
            matched_products, 
            category_tree, 
            on='Тип товара', 
            how='left'
        )

        print(f"[INFO +] Успешно проверенные  {len(validated_products)} продукты с деревом категорий.")
        return validated_products

    except Exception as e:
        print(f"[INFO -] Ошибка при проверке дерева категорий: {e}")
        return None

def main():
    
    product_list_file = "Список товаров.xlsx"
    category_tree_file = "Дерево категорий.xlsx"
    supplier_data_file = "Данные поставщика.xlsx"

    
    product_list_columns = ['Наименование', 'Тип товара']
    category_tree_columns = ['Главная категория', 'Дочерняя категория', 'Тип товара']
    supplier_data_columns = ['Наименование']

    
    product_list = read_excel_file(product_list_file, product_list_columns)
    category_tree = read_excel_file(category_tree_file, category_tree_columns)
    supplier_data = read_excel_file(supplier_data_file, supplier_data_columns)

    if product_list is None or category_tree is None or supplier_data is None:
        print("[INFO -] Ошибка: Не удалось обработать один или несколько необходимых файлов..")
        return

    
    matched_products = find_matching_products(supplier_data, product_list)

    if matched_products is not None:
        
        validated_products = validate_with_category_tree(matched_products, category_tree)

        if validated_products is not None:
            output_file = "Validated_Product_Types.xlsx"
            validated_products.to_excel(output_file, index=False)
            print(f"[INFO: SUCCESSFULLY] Validated product types saved to '{output_file}'.")

if __name__ == "__main__":
    main()
