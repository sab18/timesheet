from datetime import date

json_file_path_config = "config.json"

yaml_file_path_projects = "projects.yaml"
yaml_file_path_locations = "locations.yaml"

excel_file_path = f"Timesheet_{date.today().year}.xlsx"



config_default_data = {
    "locations_yaml": "locations.yaml", 
    "projects_yaml": "projects.yaml", 
    "timesheet_excel": "timesheet.xlsx"
}




projects_default_data= {
    "Main Projects": [
        {"project_nickname": "A1", "project_name": "Project A", "cost_center_name": "Cost Center A", "cost_center_code": "A01"},
        {"project_nickname": "B1", "project_name": "Project B", "cost_center_name": "Cost Center B", "cost_center_code": "B01"},
    ],
    "Side Projects": [
        {"project_nickname": "C1", "project_name": "Project C", "cost_center_name": "Cost Center C", "cost_center_code": "C01"},
    ],
    "Misc.": [
        {"project_nickname": "D1", "project_name": "Project D", "cost_center_name": "Cost Center D", "cost_center_code": "D01"},
    ],
    
    "Retired": [
        {"project_nickname": "E1", "project_name": "Project E", "cost_center_name": "Cost Center E", "cost_center_code": "E01"}
    ]
}


locations_default_data = ["Northborough", "Home", "GTL"]

excel_sheet_names = {"for_timesheet":"for_timesheet",
                     "daily_summaries":"daily_summaries",
                     "digest":"digest"
}

headers_default_data_for_timesheet = ["project_nickname", "project_name", "cost_center_name", "cost_center_code", "date", "hrs"]
headers_default_data_daily_summaries = ["date", "location", "note"]

