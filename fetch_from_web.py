import requests
from io import StringIO
import csv

# read a file to get what number we read last was
last_read = -1
try:
    with open("last_read_number.txt", "r") as f:
        last_read = int(f.read())
except:
    pass

result = requests.get("https://johnwesthoff.com/scoutingdata/")
fake_csv_file = StringIO(result.text)

rows = list(csv.reader(fake_csv_file))
for index, row in enumerate(rows):
  if index <= last_read:
      continue

  csv_output_filename = f'webdata/web_data_index_{index}.csv'
  with open(csv_output_filename, "w+") as new_file:
    writer = csv.writer(new_file)
    writer.writerow(row)

  last_read = index

# write to a file to save what number we read last was
try:
    with open("last_read_number.txt", "w+") as f:
        f.write(str(last_read))
except:
    pass
