import os

def get_filename_from_dir(dir_path):
    file_list = []
    for item in os.listdir(dir_path):
        basename = os.path.basename(item)
        file_list.append(basename)
    print(file_list)
    return file_list

dir = "image"
get_filename_from_dir(dir)
