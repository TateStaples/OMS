###
### Extracts the simulation file with sim extension to results_XXX folder and generates res2 files to include the location of aircraft and crew for each halo violation
###
### Usage: python3 convert.py XXX.sim 
###
import zipfile
import sys
import glob
import shutil
import os

#file_name = ""
if len(sys.argv)>1:
    # get file name from the first argument
    file_name = sys.argv[1]
    print("The arguments are: ", file_name)
else:   # print usage and quit
    print("Usage convert.py XXX.sim")
    quit()

# remove extension
file_name_only = file_name.rsplit('.', 1)[0]

#unzip files to result_XXX.sim
unzip_folder = 'result_' + file_name_only

#if unzip_folder exists, delete it
if os.path.exists(unzip_folder) and os.path.isdir(unzip_folder):
    shutil.rmtree(unzip_folder, ignore_errors=True)

with zipfile.ZipFile(file_name, 'r') as zip_ref:
    zip_ref.extractall(unzip_folder)

#read directory
res_files = glob.glob(unzip_folder + "/results/*.res")

for res_file in res_files:
    plb_file = res_file.rsplit(".", 1)[0] + ".plb"

    out_file = res_file + "2"
    with open(out_file, 'w') as outfile, open(plb_file, 'r', encoding='utf-8') as plbfile, open(res_file, 'r', encoding='utf-8') as infile:

        plbline = plbfile.readline()    # read first line of plb file
        prev_timeline = 0               # prev_timeline = 0
        prev_plbline = plbline          # just the first line
        for line in infile:
            if line.startswith("2,"):   # halo violation
                # find timeline
                halo_timeline = line.strip().split(",")[7]

                # find the timeline in plb
                while plbline:
                    plb_objs = plbline.split(";")
                    plbline_timeline = plb_objs[0]

                    if plbline_timeline == halo_timeline:
                        for obj in plb_objs:
                            if len(obj.split(",")) == 8:    # only aircraft information
                                line = line + "2-5," + obj + "\n"
                        for obj in plb_objs:
                            if len(obj.split(",")) == 5:    # only person information
                                line = line + "2-6," + obj + "\n"
                        break                        
                    elif plbline_timeline > halo_timeline:   # not exact timeline exists on the plb
                        # not all time lines are stored in plb, need to interpolate to get the expected position
                        prev_plb_objs = prev_plbline.split(";")
                        prev_ac = [];
                        prev_crew = [];
                        for obj in prev_plb_objs:
                            obj_info = obj.split(",")
                            array_length = len(obj_info)
                            if array_length == 8:           # only aircraft information
                                prev_ac.append(obj_info)
                            elif array_length == 5:         # only crew informaton
                                prev_crew.append(obj_info)

                        for obj in plb_objs:
                            ac_data = obj.split(",")
                            if len(ac_data) == 8:    # only aircraft information
                                dist = 20000
                                match_ac = None
                                for ac in prev_ac:          # find the nearest ac from previos timeline
                                                            # this usually work as airplane can move at most 3.5m for 1 sec
                                    new_dist = (float(ac[2]) - float(ac_data[2])) ** 2 + (float(ac[3]) - float(ac_data[3])) ** 2
                                    if dist > new_dist:
                                        dist = new_dist
                                        match_ac = ac                                

                                if match_ac != None:        # interpolate the timeline
                                    match_ac[2] = float(match_ac[2]) + (float(ac_data[2]) - float(match_ac[2])) * (float(plbline_timeline) - float(prev_timeline)) / (float(halo_timeline) - float(prev_timeline)) 
                                    match_ac[3] = float(match_ac[3]) + (float(ac_data[3]) - float(match_ac[3])) * (float(plbline_timeline) - float(prev_timeline)) / (float(halo_timeline) - float(prev_timeline)) 
                                    line = line + "2-5," + ",".join([str(x) for x in match_ac]) + "\n"

                        count = 0
                        for obj in plb_objs:
                            crew_data = obj.split(",")
                            if len(crew_data) == 5:    # only aircraft information
                                match_crew = prev_crew[count]
                                match_crew[2] = float(match_crew[2]) + (float(crew_data[2]) - float(match_crew[2])) * (float(plbline_timeline) - float(prev_timeline)) / (float(halo_timeline) - float(prev_timeline)) 
                                match_crew[3] = float(match_crew[3]) + (float(crew_data[3]) - float(match_crew[3])) * (float(plbline_timeline) - float(prev_timeline)) / (float(halo_timeline) - float(prev_timeline)) 
                                line = line + "2-6," + ",".join([str(x) for x in match_crew]) + "\n"
                                count+=1
                        break

                    # skip to the next line of plb
                    prev_plbline = plbline                    
                    plbline = plbfile.readline()
                    prev_timeline = plbline_timeline

            # just copy the res
            outfile.write(line)


