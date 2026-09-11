import pandas as pd
import pyodbc
import os
from datetime import datetime

# ==============================================================================
# 1. INPUT & OUTPUT PATHS
# ==============================================================================

# Input Excel file path
input_excel = r"C:\Users\PC\Desktop\Nayab's Folder\Python Data Pull Project\Input\Input_Data.xlsx"

# Output destination directory
output_folder = r"C:\Users\PC\Desktop\Nayab's Folder\Python Data Pull Project\Output"

# ==============================================================================
# 2. CURRENT RUN DATE (YYYYMMDD)
# ==============================================================================

run_date = datetime.today().strftime("%Y%m%d")

# ==============================================================================
# 3. SQL SERVER CONNECTION
# ==============================================================================

server = 'DESKTOP-VV9N5U5'
database = 'PRACTICEDB'

conn_str = (
    f"DRIVER={{ODBC Driver 17 for SQL Server}};"
    f"SERVER={server};"
    f"DATABASE={database};"
    f"Trusted_Connection=yes;"
    f"TrustServerCertificate=yes;"
)

conn = pyodbc.connect(conn_str)
cursor = conn.cursor()

# ==============================================================================
# 4. READ INPUT EXCEL WITH CONTROL TABLE LIST
# ==============================================================================

controls_df = pd.read_excel(input_excel)

# Initialize summary tracking list
summary_records = []

# ==============================================================================
# 5. PROCESS EACH CONTROL
# ==============================================================================

schema_prefix = "dbo."

for index, row in controls_df.iterrows():

    file_exported = int(row["FILE_EXPORTED"])

    # ------------------------------------------------
    # SKIP CONTROLS ALREADY EXPORTED
    # ------------------------------------------------

    if file_exported == 1:
        print(f"\nSkipping {row['REPORT_NAME']} - FILE_EXPORTED already = 1")
        continue

    control_name = str(row["REPORT_NAME"]).strip()

    pass_table = schema_prefix + str(row["PASS_TABLE"]).strip()
    fail_table = schema_prefix + str(row["FAIL_TABLE"]).strip()
    data_as_of = pd.to_datetime(row["DATA_AS_OF_DATE"]).date()

    # Extract REPORT ID from CAT_xx
    report_id = int(control_name.split("_")[1])

    print(f"\nProcessing control: {control_name}")
    print(f"  REPORT ID: {report_id}")
    print(f"  PASS TABLE: {pass_table}")
    print(f"  FAIL TABLE: {fail_table}")
    print(f"  DATA AS OF DATE: {data_as_of}")

    try:

        # ------------------------------------------------
        # PASS TABLE DATA
        # ------------------------------------------------

        query_pass = f"""
        SELECT *
        FROM {pass_table}
        """

        df_pass = pd.read_sql(query_pass, conn)

        # ------------------------------------------------
        # FAIL TABLE DATA
        # ------------------------------------------------

        query_fail = f"""
        SELECT *
        FROM {fail_table}
        """

        df_fail = pd.read_sql(query_fail, conn)

        # ------------------------------------------------
        # EXPORT EXCEL FILE
        # ------------------------------------------------

        output_file = os.path.join(
            output_folder,
            f"{control_name}_Python_{run_date}.xlsx"
        )

        with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
            df_fail.to_excel(writer, sheet_name="Fail", index=False)
            df_pass.to_excel(writer, sheet_name="Pass", index=False)

        print(f"  -> Control exported: {output_file}")

        # ------------------------------------------------
        # UPDATE SQL FILE_EXPORTED FLAG
        # ------------------------------------------------

        stats_table = schema_prefix + "REPORTS_RUN_STATS"

        update_query = f"""
        UPDATE {stats_table}
        SET FILE_EXPORTED = 1
        WHERE REPORT_ID = ?
          AND DATA_AS_OF_DATE = ?
        """

        cursor.execute(
            update_query,
            report_id,
            data_as_of
        )

        conn.commit()

        print(f"  -> SQL updated: REPORT_ID={report_id}, "
              f"DATA_AS_OF_DATE={data_as_of}, FILE_EXPORTED=1")

        # Track results for summary file
        summary_records.append({
            "REPORT_NAME": control_name,
            "REPORT_ID": report_id,
            "DATA_AS_OF_DATE": data_as_of,
            "PASS_ROWS": len(df_pass),
            "FAIL_ROWS": len(df_fail),
            "STATUS": "SUCCESS",
            "OUTPUT_FILE": output_file
        })

    except Exception as e:

        conn.rollback()

        print(f"  -> ERROR exporting {control_name}")
        print(f"     {str(e)}")

        summary_records.append({
            "REPORT_NAME": control_name,
            "REPORT_ID": report_id,
            "DATA_AS_OF_DATE": data_as_of,
            "PASS_ROWS": None,
            "FAIL_ROWS": None,
            "STATUS": "ERROR",
            "OUTPUT_FILE": str(e)
        })

# ==============================================================================
# 6. SAVE SUMMARY FILE
# ==============================================================================

summary_df = pd.DataFrame(summary_records)

summary_file = os.path.join(
    output_folder,
    "Summary.xlsx"
)

summary_df.to_excel(
    summary_file,
    index=False
)

print(
    f"\nSUMMARY FILE CREATED AND EXPORTED TO THE SAME PATH:"
    f"\n{summary_file}"
)

# ==============================================================================
# 7. CLOSE CONNECTION
# ==============================================================================

cursor.close()
conn.close()

print("\nPROCESS COMPLETED SUCCESSFULLY.")