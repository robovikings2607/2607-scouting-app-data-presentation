import requests
from io import StringIO
import csv

result = requests.get("https://johnwesthoff.com/scoutingdata/")
fake_csv_file = StringIO(result.text)

rows = list(csv.reader(fake_csv_file))
for index, row in enumerate(rows):
  csv_output_filename = f'webdata/web_data_index_{index}.csv'
  with open(csv_output_filename, "w+") as new_file:
    writer = csv.writer(new_file)
    writer.writerow(row)
