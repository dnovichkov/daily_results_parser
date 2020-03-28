import json
import datetime
import os
import xlsxwriter


def save_json(json_result, filename_prefix: str):
    json_file_name = filename_prefix + "_" + datetime.datetime.now().strftime("%Y-%m-%d %H-%M-%S") + "_.json"

    with open(json_file_name, "w") as f:
        for chunk in json.JSONEncoder(indent=4, ensure_ascii=False).iterencode(
                json_result
        ):
            f.write(chunk)


def save_results_to_excel(results, filename):
    workbook = xlsxwriter.Workbook(filename)
    worksheet = workbook.add_worksheet()
    bold = workbook.add_format({'bold': True})

    # Write some data headers.

    worksheet.set_column('A:A', 14)
    worksheet.set_column('B:B', 14)
    worksheet.set_column('C:C', 10)
    worksheet.freeze_panes(1, 0, )

    worksheet.write('A1', 'Проект', bold)
    worksheet.write('B1', 'Активность', bold)
    worksheet.write('C1', 'Длительность', bold)

    cell_format = workbook.add_format()
    row_count = 0
    for project, activities in results.items():
        for activity, duration in activities.items():
            row_count += 1
            worksheet.write(row_count, 0, project, cell_format)
            worksheet.write(row_count, 1, activity, cell_format)
            worksheet.write(row_count, 2, duration, cell_format)

    worksheet.autofilter(0, 0, len(results), 21)
    workbook.close()


def get_daily_results(text_filename: str):
    lines = []
    day_records = {}
    project_results = {}
    with open(text_filename, encoding='utf-8') as fp:
        dateline = None
        for line in fp:
            line = line.replace('\n', '')
            lines.append(line)
            if '.03' in line:
                dateline = line

                day_records[dateline] = []
            else:
                if line and dateline:
                    project_end_pos = line.find(' ')
                    project = line[:project_end_pos]
                    duration_pos = line.rfind(' ')
                    activity = line[project_end_pos + 1:duration_pos].capitalize()
                    if not activity:
                        activity = 'Работы по проекту'

                    string_duration = line[duration_pos + 1:]
                    string_duration = string_duration.replace('ч', '')
                    string_duration = string_duration.replace(',', '.')
                    duration = float(string_duration)

                    day_records[dateline].append(line)
                    if project not in project_results:
                        project_results[project] = {}
                    if activity not in project_results[project]:
                        project_results[project][activity] = 0
                    project_results[project][activity] += duration

    save_json(project_results, 'project_results')
    save_results_to_excel(project_results, 'project_results.xlsx')


if __name__ == "__main__":
    get_daily_results('result_example.txt')
