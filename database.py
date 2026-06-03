import os
import pandas as pd

# Define the name of our Excel file
EXCEL_FILE = "it_tickets_database.xlsx"


def initialize_database():
    # Check if the Excel file already exists so we don't overwrite it
    if not os.path.exists(EXCEL_FILE):
        # Define the columns for our IT ticketing system
        columns = [
            "Ticket ID",
            "Date Created",
            "Employee Name",
            "Department",
            "Category",
            "Urgency",
            "Subject",
            "Description",
            "Status",
            "IT Notes",
        ]

        # Create an empty DataFrame (table) with these columns
        df = pd.DataFrame(columns=columns)

        # Save it as an Excel file
        df.to_excel(EXCEL_FILE, index=False)
        print(f"🎉 Success! Database created as '{EXCEL_FILE}'")
    else:
        print("Database already exists. Ready to go!")


if __name__ == "__main__":
    initialize_database()