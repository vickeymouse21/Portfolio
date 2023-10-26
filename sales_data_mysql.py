import pandas as pd
import pymysql

# Buat koneksi ke MySQL
conn = pymysql.connect(
    host='localhost',
    user='your_username',
    password='your_password',
    database='your_database'
)

# Baca data dari Excel
daily_unit_sales_df = pd.read_excel('sales_data.xlsx', sheet_name='daily_unit_sales')
product_tbl_df = pd.read_excel('sales_data.xlsx', sheet_name='product_tbl')

# Simpan data ke MySQL
daily_unit_sales_df.to_sql('daily_unit_sales', conn, if_exists='replace', index=False)
product_tbl_df.to_sql('product_tbl', conn, if_exists='replace', index=False)

# Commit perubahan
conn.commit()

# Income Statement (query diubah menjadi MySQL)
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
            Product, Quantity, Revenue, Profit, FORMAT(100.0*Profit/Revenue, 2) AS GPM
        FROM new_data;
    """, conn
)

# SALES QUANTITY REPORT (query diubah menjadi MySQL)
penjualan_tiap_periode = pd.read_sql_query(
    """
        SELECT
            DATE_FORMAT(dus.day_date, '%Y-%m') AS Periode,
            SUM(dus.units_sold) AS Quantity
        FROM daily_unit_sales dus
        GROUP BY Periode
        ORDER BY Periode;
    """, conn
)

# Total Penjualan (query diubah menjadi MySQL)
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

# Top Sales Performance (query diubah menjadi MySQL)
top_sales_performance = pd.read_sql_query(
    """
        WITH new_data AS (
            SELECT
                DATE_FORMAT(day_date, '%Y-%m') AS Periode,
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

# Tutup koneksi
conn.commit()
conn.close()
