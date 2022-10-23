import shutil
import os

'''
The function move file to different folder with keyword selection feature
parent_dir: the path that contains the folder
folder_name: the folder that will move all files inside it
new_path: the location for the new folder
keyword: input '' if no need for keyword, input anyother keyword as needed as a string
'''
def file_merge(parent_dir, folder_name, new_path, keyword):
    # Create new folder and combine info of old folder
    os.mkdir(new_path)
    old_path = os.path.join(parent_dir, folder_name)

    for (root, dirs, files) in os.walk(old_path):
        for names in files:
            if keyword in names:
                file_name = os.path.join(root, names)

                # find the folder that contain the file and combine it with file name
                folder_path = root.replace(parent_dir, '')
                folder_path = folder_path[1:]
                folder_path = folder_path.replace('\\', '_')
                combine_name = folder_path + " " + names

                new_place = os.path.join(new_path, combine_name)
                shutil.copy(file_name, new_place)




if __name__ == "__main__":
    parent_dir = 'C:\\Users\\10648\\OneDrive\\NCPR Intern\\File info gather'
    folder_name = 'Demo1'
    new_path = 'C:\\Users\\10648\\OneDrive\\NCPR Intern\\File info gather\\New folder'
    keyword = ''
    file_merge(parent_dir, folder_name, new_path, keyword)
