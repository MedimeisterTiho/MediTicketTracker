import sys
import os
import mechanicalsoup
import pandas as pd
import datetime
import multiprocessing
from multiprocessing import Pool, freeze_support


def load_data(file_path, data_sheet, codes_sheet):
    """Load Excel sheets into Pandas DataFrames."""
    data_df = pd.read_excel(file_path, sheet_name=data_sheet)
    codes_df = pd.DataFrame(columns=["ID", "Vorname", "Nachname", "Email", "Code", "Status"])  # Clear the codes sheet
    return data_df, codes_df


def process_single_row(args):
    """Process a single row with a new browser instance."""
    row, base_url, exp_url = args
    browser = mechanicalsoup.StatefulBrowser()
    browser.open(base_url)
    browser.select_form('form[action="https://ticket.medimeisterschaften.com/redeem"]')

    if row.get("Status") == "eingelöst":
        return row, None  # Already redeemed, skip processing

    ticket_code = row.get("Code")
    if pd.isna(ticket_code):
        return row, None  # No code, skip processing

    browser["voucher"] = ticket_code
    response = browser.submit_selected()

    if exp_url in response.url:
        row["Status"] = "eingelöst"
        print(f"{ticket_code} eingelöst")
        return row, None  # Successfully redeemed
    else:
        row["Status"] = "nicht eingelöst"
        print(f"{ticket_code} nicht eingelöst")
        return row, {
            "ID": row.get("ID"),
            "Vorname": row.get("Vorname"),
            "Nachname": row.get("Nachname"),
            "Email": row.get("Email"),
            "Code": ticket_code,
            "Status": "nicht eingelöst",
        }


def process_vouchers_parallel(data_df, codes_df, base_url, exp_url):
    """Process voucher codes in parallel using multiprocessing."""

    with Pool(processes=multiprocessing.cpu_count()) as pool:
        # Prepare arguments for processing
        args = [(row, base_url, exp_url) for _, row in data_df.iterrows()]

        # Process rows in parallel
        results = pool.map(process_single_row, args)

        processed_rows = []
        new_entries = []

        # Collect results
        for row, entry in results:
            processed_rows.append(row)
            if entry:
                new_entries.append(entry)

        # Update dataframes
        data_df = pd.DataFrame(processed_rows)
        if new_entries:
            new_codes_df = pd.DataFrame(new_entries)
            codes_df = pd.concat([codes_df, new_codes_df], ignore_index=True)

        return data_df, codes_df


def save_data(file_path, data_df, codes_df, data_sheet, codes_sheet):
    """Save the updated data frames to the Excel file."""
    with pd.ExcelWriter(file_path, engine="xlsxwriter") as writer:
        data_df.to_excel(writer, sheet_name=data_sheet, index=False)
        codes_df.to_excel(writer, sheet_name=codes_sheet, index=False)


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


def main():
    # Configurations
    workbook_file = "Ticketcodes.xlsx"
    data_sheet_name = "Daten"
    codes_sheet_name = "Freie Codes"
    base_url = "https://ticket.medimeisterschaften.com/"
    exp_url = "https://ticket.medimeisterschaften.com/?voucher_invalid"
    
    cwd = os.path.dirname(os.path.abspath(__file__))
    workbook_file = os.path.join("./", workbook_file)
    print(workbook_file)
    print(f"Ticket Tracker - Medimeisterschaften {datetime.datetime.now().year}")

    # Load data
    data_df, codes_df = load_data(workbook_file, data_sheet_name, codes_sheet_name)
    assert not data_df.empty, "Data sheet not found"

    # Process voucher codes in parallel
    try:
        data_df, codes_df = process_vouchers_parallel(data_df, codes_df, base_url, exp_url)
    except Exception as e:
        print(f"An error occurred: {e}")
        sys.exit(1)
    finally:
        # Save updated data
        save_data(workbook_file, data_df, codes_df, data_sheet_name, codes_sheet_name)


if __name__ == "__main__":
    freeze_support()
    main()
