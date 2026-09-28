import sqlite3
import pandas as pd
import matplotlib.pyplot as plt


# ============================================================
# 1. CONNECT TO DATABASE
# ============================================================

conn = sqlite3.connect("store.db")


# ============================================================
# 2. MONTHLY QUANTITY
# ============================================================

query1 = """
SELECT
    strftime('%Y-%m', sale_date) AS month,
    SUM(quantity) AS total_quantity
FROM sales
GROUP BY month
ORDER BY month
"""

monthly_quantity_df = pd.read_sql_query(query1, conn)


# ============================================================
# 3. MONTHLY REVENUE
# ============================================================

query2 = """
SELECT
    strftime('%Y-%m', s.sale_date) AS month,
    SUM(s.quantity * p.price) AS total_revenue
FROM sales s
JOIN products p
    ON s.product_id = p.product_id
GROUP BY month
ORDER BY month
"""

monthly_revenue_df = pd.read_sql_query(query2, conn)


# ============================================================
# 4. CATEGORY REVENUE
# ============================================================

query3 = """
SELECT
    CASE
        WHEN p.product_name = 'Laptop'
            THEN 'Computers'
        WHEN p.product_name IN
            ('Mouse', 'Keyboard', 'Headphones', 'Webcam', 'USB Cable')
            THEN 'Accessories'
        WHEN p.product_name IN
            ('Monitor', 'Printer')
            THEN 'Office Equipment'
        ELSE 'Other'
    END AS category,

    SUM(s.quantity * p.price) AS total_revenue

FROM sales s

JOIN products p
    ON s.product_id = p.product_id

GROUP BY category
ORDER BY total_revenue DESC
"""

category_revenue_df = pd.read_sql_query(query3, conn)


# ============================================================
# 5. CITY REVENUE
# ============================================================

query4 = """
SELECT
    city,
    SUM(s.quantity * p.price) AS total_revenue
FROM sales s
JOIN products p
    ON s.product_id = p.product_id
GROUP BY city
ORDER BY total_revenue DESC
"""

city_revenue_df = pd.read_sql_query(query4, conn)


# ============================================================
# 6. TOP PRODUCT BY CITY
# ============================================================

query5 = """
WITH product_city_revenue AS (

    SELECT
        s.city,
        p.product_name,
        SUM(s.quantity * p.price) AS revenue

    FROM sales s

    JOIN products p
        ON s.product_id = p.product_id

    GROUP BY
        s.city,
        p.product_name
),

ranked_products AS (

    SELECT
        city,
        product_name,
        revenue,

        ROW_NUMBER() OVER (
            PARTITION BY city
            ORDER BY revenue DESC
        ) AS rank_number

    FROM product_city_revenue
)

SELECT
    city,
    product_name,
    revenue
FROM ranked_products
WHERE rank_number = 1
ORDER BY city
"""

top_product_by_city_df = pd.read_sql_query(query5, conn)


# ============================================================
# 7. AVERAGE SALE VALUE BY CATEGORY
# ============================================================

query6 = """
SELECT
    CASE
        WHEN p.product_name = 'Laptop'
            THEN 'Computers'
        WHEN p.product_name IN
            ('Mouse', 'Keyboard', 'Headphones', 'Webcam', 'USB Cable')
            THEN 'Accessories'
        WHEN p.product_name IN
            ('Monitor', 'Printer')
            THEN 'Office Equipment'
        ELSE 'Other'
    END AS category,

    AVG(s.quantity * p.price) AS average_sale_value

FROM sales s

JOIN products p
    ON s.product_id = p.product_id

GROUP BY category
ORDER BY average_sale_value DESC
"""

average_sale_category_df = pd.read_sql_query(query6, conn)


# ============================================================
# 8. PRINT RESULTS
# ============================================================

print("\n" + "=" * 60)
print("MONTHLY QUANTITY")
print("=" * 60)
print(monthly_quantity_df)


print("\n" + "=" * 60)
print("MONTHLY REVENUE")
print("=" * 60)
print(monthly_revenue_df)


print("\n" + "=" * 60)
print("CATEGORY REVENUE")
print("=" * 60)
print(category_revenue_df)


print("\n" + "=" * 60)
print("CITY REVENUE")
print("=" * 60)
print(city_revenue_df)


print("\n" + "=" * 60)
print("TOP PRODUCT BY CITY")
print("=" * 60)
print(top_product_by_city_df)


print("\n" + "=" * 60)
print("AVERAGE SALE VALUE BY CATEGORY")
print("=" * 60)
print(average_sale_category_df)


# ============================================================
# 9. LINE CHART - MONTHLY QUANTITY
# ============================================================

plt.figure(figsize=(10, 5))

plt.plot(
    monthly_quantity_df["month"],
    monthly_quantity_df["total_quantity"],
    marker="o"
)

plt.title("Monthly Sales Quantity")
plt.xlabel("Month")
plt.ylabel("Total Quantity")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig("monthly_quantity.png")
plt.show()


# ============================================================
# 10. LINE CHART - MONTHLY REVENUE
# ============================================================

plt.figure(figsize=(10, 5))

plt.plot(
    monthly_revenue_df["month"],
    monthly_revenue_df["total_revenue"],
    marker="o"
)

plt.title("Monthly Revenue")
plt.xlabel("Month")
plt.ylabel("Total Revenue")
plt.xticks(rotation=45)
plt.tight_layout()

plt.savefig("monthly_revenue.png")
plt.show()


# ============================================================
# 11. CREATE SUMMARY
# ============================================================

total_revenue = monthly_revenue_df["total_revenue"].sum()

highest_month_row = monthly_revenue_df.loc[
    monthly_revenue_df["total_revenue"].idxmax()
]

highest_category_row = category_revenue_df.loc[
    category_revenue_df["total_revenue"].idxmax()
]

highest_city_row = city_revenue_df.loc[
    city_revenue_df["total_revenue"].idxmax()
]

summary_df = pd.DataFrame({
    "Metric": [
        "Total Revenue",
        "Highest Revenue Month",
        "Highest Revenue Category",
        "Highest Revenue City"
    ],

    "Value": [
        total_revenue,
        highest_month_row["month"],
        highest_category_row["category"],
        highest_city_row["city"]
    ]
})


# ============================================================
# 12. EXPORT EVERYTHING TO EXCEL
# ============================================================

with pd.ExcelWriter(
    "sales_report.xlsx",
    engine="openpyxl"
) as writer:

    summary_df.to_excel(
        writer,
        sheet_name="Summary",
        index=False
    )

    monthly_quantity_df.to_excel(
        writer,
        sheet_name="Monthly Quantity",
        index=False
    )

    monthly_revenue_df.to_excel(
        writer,
        sheet_name="Monthly Revenue",
        index=False
    )

    category_revenue_df.to_excel(
        writer,
        sheet_name="Category Revenue",
        index=False
    )

    city_revenue_df.to_excel(
        writer,
        sheet_name="City Revenue",
        index=False
    )

    top_product_by_city_df.to_excel(
        writer,
        sheet_name="Top Product by City",
        index=False
    )

    average_sale_category_df.to_excel(
        writer,
        sheet_name="Average Sale Category",
        index=False
    )


# ============================================================
# 13. CLOSE DATABASE
# ============================================================

conn.close()


print("\n" + "=" * 60)
print("REPORT GENERATED SUCCESSFULLY!")
print("=" * 60)
print("Created files:")
print("- sales_report.xlsx")
print("- monthly_quantity.png")
print("- monthly_revenue.png")