import json
import csv

# Open the JSON file and load its contents into a Python object
with open('files/instance-info.json') as f:
    data = json.load(f)

# Access the data
instance_info = data['InstanceInformationList']

header = ['Instance ID', 'Computer Name', 'PingStatus', 'Version']

# Open the CSV file in write mode and specify newline='' to avoid extra blank rows
with open('files/my_file.csv', mode='w', newline='') as file:

    # Create a CSV writer object
    writer = csv.writer(file)
    writer.writerow(header)

    # Write the header row to the CSV file
    for instance_info2 in instance_info:

        row = [instance_info2['InstanceId'], instance_info2['ComputerName'], instance_info2['PingStatus'], instance_info2['AgentVersion'] ]

        writer.writerow(row)

