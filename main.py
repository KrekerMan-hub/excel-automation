#===================
# Main file
#===================

from pathlib import Path

import pandas as pd


BASE_DIR = Path(__file__).resolve().parent
INPUT_PATH = BASE_DIR/"example-input"/"orders.csv"
OUTPUT_DIR = BASE_DIR/"example_output"
OUTPUT_PATH = OUTPUT_DIR/"orders_report.xlsx"

def load_orders(input_path: Path) -> pd.DataFrame:
    df = pd.read_csv(input_path, sep=";", encoding="utf-8")

    required_columns = {"order_id", "customer", "price", "quantity"}
    missing_columns = required_columns - set(df.columns)

    if missing_columns:
        names = ", ".join(sorted(missing_columns))
        raise ValueError(f"В файле отсутствуют столбцы: {names}")

    return df

def process_orders(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, int]:
    df = df.copy()
    # Duplicates_delete
    orders_before = df.shape[0]
    df["customer"] = df["customer"].str.strip()
    df = df.drop_duplicates()
    orders_after = df.shape[0]
    deleted_rows = orders_before - orders_after

    # Price_check
    df["original_price"] = df["price"]
    df["price"] = pd.to_numeric(df["price"], errors="coerce")
    invalid_price = df["price"].isna() | (df["price"] <= 0)

    # Quantity_check
    df["original_quantity"] = df["quantity"]
    df["quantity"] = pd.to_numeric(df["quantity"], errors="coerce")
    invalid_quantity = df["quantity"].isna() | (df["quantity"] <= 0) | (df["quantity"] % 1 != 0)
    
    # Error_reason
    df["error_reason"] = ""
    df.loc[invalid_price, "error_reason"] = ("Цена должна быть числом больше 0; ")
    df.loc[invalid_quantity, "error_reason"] += ("Количество должно быть положительным целым числом; ")

    # problems
    invalid_orders = invalid_price | invalid_quantity
    problem_orders = df[invalid_orders].copy()
    valid_orders = df[~invalid_orders].copy()

    # Total_price
    valid_orders["total"] = (valid_orders["price"] * valid_orders["quantity"]).round(2)

    return valid_orders, problem_orders, deleted_rows

def save_report(
    valid_orders: pd.DataFrame, problem_orders: pd.DataFrame, output_path: Path,) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    
    # Output Excel
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        valid_orders.to_excel(
            writer,
            sheet_name="Корректные заказы",
            index=False,
        )
        problem_orders.to_excel(
            writer,
            sheet_name="Проблемные заказы",
            index=False
        )


def main() -> None:
    try:
        df = load_orders(INPUT_PATH)

        valid_orders, problem_orders, deleted_rows = process_orders(df)

        save_report(valid_orders, problem_orders, OUTPUT_PATH)
    
    except FileNotFoundError:
        print(f"Исходный файл не найден: {INPUT_PATH}")
        return

    except PermissionError:
        print(
            "Нет доступа к файлу. "
            "Если отчёт открыт в Excel, закрой его и повтори запуск."
        )
        return
    
    except ValueError as error:
        print(f"Не удалось обработать данные: {error}")
        return

    total_amount = valid_orders["total"].sum()

    print(f"Исходных строк: {df.shape[0]}")
    print(f"Удалено дублей: {deleted_rows}")
    print(f"Корректных заказов: {valid_orders.shape[0]}")
    print(f"Проблемных заказов: {problem_orders.shape[0]}")
    print(f"Общая стоимость: {total_amount:.2f}")
    print(f"Отчёт сохранён: {OUTPUT_PATH}")


if __name__ == "__main__":
    main()