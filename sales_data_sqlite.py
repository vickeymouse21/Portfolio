import sqlite3
import pandas as pd

conn = sqlite3.connect('portfolio.db')

daily_unit_sales_df = pd.read_excel('sales_data.xlsx', sheet_name='daily_unit_sales')
product_tbl_df = pd.read_excel('sales_data.xlsx', sheet_name='product_tbl')

daily_unit_sales_df.to_sql('daily_unit_sales', conn, if_exists='replace', index=False)
product_tbl_df.to_sql('product_tbl', conn, if_exists='replace', index=False)

conn.commit()

# Income Statement
income_statement = pd.read_sql_query(
    """
        WITH new_data AS (
            SELECT
                dus.product_name AS Product,
                SUM(dus.units_sold) AS Quantity,
                SUM(dus.units_sold*pt.unit_price) AS Revenue,
                SUM((dus.units_sold*pt.unit_price) - (dus.units_sold*pt.unit_cost)) AS Profit
            FROM daily_unit_sales dus INNER JOIN product_tbl pt
                ON dus.product_name = pt.product_name
            GROUP BY dus.product_name
        )
        SELECT
            Product, Quantity, Revenue, Profit, printf("%.2f", (Profit*100)/Revenue) AS "GPM(%)"
        FROM new_data;
    """, conn
)

# SALES QUANTITY REPORT
# Total Penjualan tiap Periode
penjualan_tiap_periode = pd.read_sql_query(
    """
        SELECT
            strftime('%Y-%m', dus.day_date) AS Periode,
            SUM(dus.units_sold) AS Quantity
        FROM daily_unit_sales dus
        GROUP BY Periode
        ORDER BY Periode;
    """, conn
)

# Total Penjualan
total_penjualan = pd.read_sql_query(
    """
        SELECT
            SUM(dus.units_sold) AS "Total Quantity",
            SUM(dus.units_sold*pt.unit_price) AS "Total Revenue",
            SUM((dus.units_sold*pt.unit_price) - (dus.units_sold*pt.unit_cost)) AS "Total Profit"
        FROM daily_unit_sales dus INNER JOIN product_tbl pt
            ON dus.product_name = pt.product_name;
    """, conn
)

# Top Sales Performance
top_sales_performance = pd.read_sql_query(
    """
        WITH new_data AS (
            SELECT
                strftime('%Y-%m', day_date) AS Periode,
                SUM(dus.units_sold) AS TotalQuantity,
                SUM(dus.units_sold*pt.unit_price) AS TotalRevenue,
                SUM((dus.units_sold*pt.unit_price) - (dus.units_sold*pt.unit_cost)) AS TotalProfit
                FROM daily_unit_sales dus INNER JOIN product_tbl pt
                    ON dus.product_name = pt.product_name
                GROUP BY Periode
        )

        SELECT
            Periode AS "Top Periode",
            MAX(TotalQuantity) AS "Top Quantity",
            MAX(TotalRevenue) AS "Top Revenue",
            MAX(TotalProfit) AS "Top Profit"
        FROM new_data;
    """, conn
)

print("Income Statement")
print(income_statement)
print("\n")

print("Penjualan Tiap Periode")
print(penjualan_tiap_periode)
print("\n")

print("Total Penjualan")
print(total_penjualan)
print("\n")

print("Top Sales Performance")
print(top_sales_performance)

conn.commit()
conn.close()