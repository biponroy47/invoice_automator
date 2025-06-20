import csv

data = [
    {'Name': 'Alice', 'Age': 30, 'City': 'New York'},
    {'Name': 'Bob', 'Age': 25, 'City': 'Los Angeles'},
    {'Name': 'Charlie', 'Age': 35, 'City': 'Chicago'}
]

fieldnames = ['Name', 'Age', 'City']

with open('output.csv', 'w', newline='') as csvfile:
    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
    writer.writeheader()
    writer.writerows(data)