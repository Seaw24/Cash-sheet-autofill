if __name__ == "__main__":
    # Run the processor
    weekday = None
    while not weekday:
        report_date = input("What cash sheet day that you trynna fill in? ")
        # Default is yesterday
        if report_date == "":
            yesterday = date.today() - timedelta(days=1)
            report_date = yesterday.strftime("%m/%d/%Y")
        weekday = get_weekday_name(report_date)

    execute(REPORTS_FOLDER, CASH_SHEET_FOLDER,
            weekday=weekday, report_date=report_date)