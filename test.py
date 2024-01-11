import json

with open("settings.json") as jsonfile:
    config = json.load(jsonfile)
departments_dict = config.get('default_pass', '')

print(departments_dict)

