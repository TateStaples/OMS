"""
Programmatically creates a SetupBatch.json to run in the OMS

mess with the params below the imports
*requires Setup.json to reference
"""


import json
from copy import deepcopy

aircraft_counts = [(13, 2), (17, 3),  (20, 5)]
multipliers = [.25, .5, .75, .9, 1.1, 1.25, 1.5, 1.75]
variable_names = ["SlipTripFallTime"]
trials = 30

json_file = "Setup.json"
save_name = "SetupBatch.json"
file_names = "names.txt"
baseline = dict(json.load(open(json_file)))

name_file = open(file_names, 'w')

jsons = list()
for f18, e2 in aircraft_counts:
    count = f18 + e2
    for mult in multipliers:
        for variable in variable_names:
            preferences = deepcopy(baseline)
            variable_settings = preferences[variable]
            previous_mean = variable_settings["Mean"]
            print(previous_mean, mult, previous_mean * mult)
            variable_settings["Mean"] = previous_mean * mult

            print((mult-1) * 100)
            adjusted = round((mult-1) * 100)
            if mult != 1:
                name = variable + str(adjusted) + "(" + str(f18+e2) + ")"
            else:
                name = "Base(" + str(count) + ")"
            print(name)
            name_file.write(name + '\n')
            new_json = {'n': trials, 'F18': f18, "E2": e2, variable: variable_settings}
            jsons.append(new_json)
with open(save_name, 'w') as file:
    file.write('[\n')
    for i, config in enumerate(jsons):
        config_str = str(config)
        # print(config_str)
        config_str = config_str.replace("'", '"')
        config_str = config_str.replace("True", "true")
        config_str = config_str.replace("False", "false")
        # print(config_str)
        ending = '\n' if i == len(jsons)-1 else ',\n'
        file.write(config_str + ending)
    file.write(']')
# json.dump(jsons, open(save_name, 'w'))


if __name__ == '__main__':
    pass