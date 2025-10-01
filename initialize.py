import os
import yaml
import xlsxwriter
import json

from inputs import (json_file_path_config,
                    config_default_data,

                    yaml_file_path_projects,
                    projects_default_data,

                    yaml_file_path_locations,
                    locations_default_data,

                    excel_file_path,
                    excel_sheet_names,
                    headers_default_data_for_timesheet,
                    headers_default_data_daily_summaries)

def initialize_json(path, data):
    if not os.path.exists(path):
        with open(path, "w") as file:
            json.dump(data, file, indent=4)
            print(f"New json file created: {path}")

def initialize_yaml(path, data):
    if not os.path.exists(path):
        with open(path, "w") as file:
            yaml.dump(data, file)
            print(f"New yaml file created: {path}")
    else:
        print(f"yaml already exists: {path}")
        


def initialize_excel(path):
    if not os.path.exists(path):
        workbook = xlsxwriter.Workbook(path)
        worksheet_summary = workbook.add_worksheet(excel_sheet_names["digest"])
        

        worksheet_timesheet = workbook.add_worksheet(excel_sheet_names["for_timesheet"])
        for col, header in enumerate(headers_default_data_for_timesheet):
            worksheet_timesheet.write(0, col, header)

        worksheet_daily_summaries = workbook.add_worksheet(excel_sheet_names["daily_summaries"])
        for col, header in enumerate(headers_default_data_daily_summaries):
            worksheet_daily_summaries.write(0, col, header)

        workbook.close()
        print(f"New excel file created: {path}")
    else:
        print(f"excel already exists: {path}")


if __name__ == "__main__":
    initialize_json(json_file_path_config, config_default_data)
    initialize_yaml(yaml_file_path_projects, projects_default_data)
    initialize_yaml(yaml_file_path_locations, locations_default_data)
    initialize_excel(excel_file_path)
    