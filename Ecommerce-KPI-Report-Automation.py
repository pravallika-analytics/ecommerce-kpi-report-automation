import pandas as pd
from pathlib import Path

# ----------------------------
# CONFIG
# ----------------------------
INPUT_FILE = Path("C:/Users/prava/Downloads/archive/ecommerce_customer_data_custom_ratios.csv")   # put CSV in same folder as this script OR update path
OUTPUT_FILE = Path("Ecommerce_KPI_Report.xlsx")         # MUST end with .xlsx


def load_data(filepath: Path) -> pd.DataFrame:
    df = pd.read_csv(filepath)

    # Parse dates
    df["Purchase Date"] = pd.to_datetime(df["Purchase Date"], errors="coerce")

    # Clean numeric columns
    df["Total Purchase Amount"] = pd.to_numeric(df["Total Purchase Amount"], errors="coerce")
    df["Quantity"] = pd.to_numeric(df["Quantity"], errors="coerce")

    # Returns might be Yes/No or 0/1 — normalize to 0/1
    if df["Returns"].dtype == object:
        df["Returns"] = df["Returns"].astype(str).str.strip().str.lower().map(
            {"yes": 1, "y": 1, "true": 1, "1": 1, "no": 0, "n": 0, "false": 0, "0": 0}
        )
    df["Returns"] = pd.to_numeric(df["Returns"], errors="coerce").fillna(0).astype(int)

    # Churn may be Yes/No or 0/1 — normalize to 0/1
    if df["Churn"].dtype == object:
        df["Churn"] = df["Churn"].astype(str).str.strip().str.lower().map(
            {"yes": 1, "y": 1, "true": 1, "1": 1, "no": 0, "n": 0, "false": 0, "0": 0}
        )
    df["Churn"] = pd.to_numeric(df["Churn"], errors="coerce").fillna(0).astype(int)

    # Drop rows without a valid date or amount
    df = df.dropna(subset=["Purchase Date", "Total Purchase Amount"])

    return df


def build_kpis(df: pd.DataFrame):
    # Daily KPIs
    daily = (
        df.assign(Date=df["Purchase Date"].dt.date)
          .groupby("Date")
          .agg(
              Revenue=("Total Purchase Amount", "sum"),
              Orders=("Customer ID", "count"),
              Units=("Quantity", "sum"),
              Returning_Customers=("Customer ID", "nunique"),
              Return_Rate=("Returns", "mean"),
              Churn_Rate=("Churn", "mean"),
          )
          .reset_index()
    )
    daily["AOV"] = daily["Revenue"] / daily["Orders"]

    # Monthly KPIs
    monthly = (
        df.assign(Month=df["Purchase Date"].dt.to_period("M").astype(str))
          .groupby("Month")
          .agg(
              Revenue=("Total Purchase Amount", "sum"),
              Orders=("Customer ID", "count"),
              Units=("Quantity", "sum"),
              Unique_Customers=("Customer ID", "nunique"),
              Return_Rate=("Returns", "mean"),
              Churn_Rate=("Churn", "mean"),
          )
          .reset_index()
    )
    monthly["AOV"] = monthly["Revenue"] / monthly["Orders"]

    # Category KPIs
    category = (
        df.groupby("Product Category")
          .agg(
              Revenue=("Total Purchase Amount", "sum"),
              Orders=("Customer ID", "count"),
              Units=("Quantity", "sum"),
              Unique_Customers=("Customer ID", "nunique"),
              Return_Rate=("Returns", "mean"),
              Churn_Rate=("Churn", "mean"),
          )
          .reset_index()
          .sort_values("Revenue", ascending=False)
    )
    category["AOV"] = category["Revenue"] / category["Orders"]

    # Summary KPIs
    summary = pd.DataFrame(
        {
            "Metric": [
                "Total Revenue",
                "Total Orders",
                "Total Units",
                "Unique Customers",
                "Average Order Value (AOV)",
                "Return Rate",
                "Churn Rate",
                "Date Range",
            ],
            "Value": [
                df["Total Purchase Amount"].sum(),
                len(df),
                df["Quantity"].sum(),
                df["Customer ID"].nunique(),
                df["Total Purchase Amount"].sum() / max(len(df), 1),
                df["Returns"].mean(),
                df["Churn"].mean(),
                f"{df['Purchase Date'].min().date()} to {df['Purchase Date'].max().date()}",
            ],
        }
    )

    return summary, daily, monthly, category


def export_excel(summary, daily, monthly, category, output_path: Path):
    output_path.parent.mkdir(parents=True, exist_ok=True)

    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
        summary.to_excel(writer, sheet_name="Summary", index=False)
        daily.to_excel(writer, sheet_name="Daily_KPIs", index=False)
        monthly.to_excel(writer, sheet_name="Monthly_KPIs", index=False)
        category.to_excel(writer, sheet_name="Category_KPIs", index=False)

        workbook = writer.book

        money_fmt = workbook.add_format({"num_format": "$#,##0.00"})
        int_fmt = workbook.add_format({"num_format": "0"})
        pct_fmt = workbook.add_format({"num_format": "0.00%"})
        header_fmt = workbook.add_format({"bold": True})

        def format_sheet(sheet_name, money_cols=None, int_cols=None, pct_cols=None, col_width=18):
            ws = writer.sheets[sheet_name]

            # Make headers bold + set column widths
            ws.set_row(0, None, header_fmt)
            ws.set_column(0, 0, 14)  # first column slightly narrower (Date/Month/Category)
            ws.set_column(1, 50, col_width)

            if money_cols:
                for c in money_cols:
                    ws.set_column(c, c, col_width, money_fmt)
            if int_cols:
                for c in int_cols:
                    ws.set_column(c, c, col_width, int_fmt)
            if pct_cols:
                for c in pct_cols:
                    ws.set_column(c, c, col_width, pct_fmt)

        # Summary formatting: value column format depends on row; keep simple
        format_sheet("Summary")

        # Daily/Monthly/Category formatting
        # columns order:
        # Date/Month/Category | Revenue | Orders | Units | Unique/Returning Customers | Return_Rate | Churn_Rate | AOV
        format_sheet("Daily_KPIs", money_cols=[1, 7], int_cols=[2, 3, 4], pct_cols=[5, 6])
        format_sheet("Monthly_KPIs", money_cols=[1, 7], int_cols=[2, 3, 4], pct_cols=[5, 6])
        format_sheet("Category_KPIs", money_cols=[1, 7], int_cols=[2, 3, 4], pct_cols=[5, 6])

    print(f"✅ Report created: {output_path.resolve()}")


def main():
    if not INPUT_FILE.exists():
        raise FileNotFoundError(
            f"Cannot find {INPUT_FILE}. Put ecommerce_customer_data_large.csv in the same folder as this script "
            f"or update INPUT_FILE path."
        )

    df = load_data(INPUT_FILE)
    summary, daily, monthly, category = build_kpis(df)
    export_excel(summary, daily, monthly, category, OUTPUT_FILE)


if __name__ == "__main__":
    main()
