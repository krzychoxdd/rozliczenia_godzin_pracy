import pandas as pd
import datetime


def get_next_task_start(tasks, curr_i):
    first_task_start = tasks[0].split(maxsplit=1)[0]
    end_day_time = datetime.datetime.strptime(
        first_task_start, "%H:%M") + datetime.timedelta(hours=8)
    end_day_time = end_day_time.strftime("%H:%M")

    lenTasks = len(tasks)

    if curr_i + 1 < lenTasks:
        task_next = tasks[curr_i+1]
        task_start = task_next.split(maxsplit=1)[0]
    else:
        task_start = end_day_time

    return task_start


m = "marzec2023"

with open('source/'+str(m)+'.txt', 'r') as f:
    data = [line.rstrip() for line in f.readlines()]

df = pd.DataFrame(columns=['Data', 'Zadanie', 'Start', 'Koniec', 'Ilość'])
first_task_rows = []

for num, line in enumerate(data, 1):
    line_date = line.split("|")[0]
    tasks_all = line.split("|")[1]
    tasks_day_list = tasks_all.split(";")
    tasks_day_list = list(filter(bool, tasks_day_list))
    tasks_day_list = [task.strip() for task in tasks_day_list]
    prev_task_end = ''
    first_task_rows.append(len(df)+1)

    for num, task in enumerate(tasks_day_list, 1):
        if task.replace(" ", ""):
            next_task_start = get_next_task_start(tasks_day_list, num-1)
            task_start = task.split(maxsplit=1)[0]
            task_name = task.split(maxsplit=1)[1]
            if next_task_start:
                time_diff = pd.to_datetime(
                    next_task_start) - pd.to_datetime(task_start)
            else:
                time_diff = 0

            if time_diff == 0:
                total_seconds = 0
            else:
                total_seconds = time_diff.total_seconds()

            time_hours = round(total_seconds / 3600, 2)

            prev_task_end = task_start
            prev_task_name = task_name
            df.loc[len(df)] = [line_date, prev_task_name,
                               task_start, next_task_start, time_hours]

writer = pd.ExcelWriter('dest/'+str(m)+'.xlsx', engine='xlsxwriter')
df.to_excel(writer, sheet_name='Sheet1', index=False)
workbook = writer.book
worksheet = writer.sheets['Sheet1']

column_bold_text = workbook.add_format({'bold': True})

worksheet.set_column('A:A', 25)
worksheet.set_column('B:B', 120)
worksheet.set_column('C:D', 20)
worksheet.set_column('E:E', 20)

workbook = writer.book
worksheet = writer.sheets['Sheet1']
bold_border = workbook.add_format()
bold_border.set_top(1)

prev_date = ''
for row in first_task_rows:
    worksheet.set_row(row, None, bold_border)


num_days = len(df['Data'].unique())

total_hours = df['Ilość'].sum()

summary_format = workbook.add_format({'bold': True, 'align': 'right'})

worksheet.write(len(df)+3, 3, "Suma przepracowanych dni:", summary_format)
worksheet.write(len(df)+4, 3, "Suma przepracowanych godzin:", summary_format)
worksheet.write(len(df)+3, 4, num_days, summary_format)
worksheet.write(len(df)+4, 4, total_hours, summary_format)

writer.save()
