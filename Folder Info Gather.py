import os
import pandas as pd


# The function that collects folder's size
# path: the path to the folder
def get_dir_size(path):
    total = 0
    with os.scandir(path) as it:
        for entry in it:
            if entry.is_file():
                total += entry.stat().st_size
            elif entry.is_dir():
                total += get_dir_size(entry.path)
    return total


# Append multiple values to a key in the given dictionary
# dict: the dict to append
# key: header of list
# list_of_values: value to append
def add_values_in_dict(dict, key, list_of_values):
    if key not in dict:
        dict[key] = list()
    dict[key].extend(list_of_values)
    return dict


# Append multiple values to a key in the given dictionary
# dictset: the dict to append
# key_set: header of list
# listset: value to append
def append_dic(dictset, key_set, listset):
    for i in range(len(key_set)):
        dictset[key_set[i]].append(listset[i])
    return dictset


# Collect info of a folder
# path: path to the folder
# folder_dict: dict to collect info
# keyset: dict's header
# check_size: the size redline
def collect_info(path, folder_dict, keyset, check_size):
    # Create string and data frame to load information of folders
    total_information = ''

    # Iterate through files
    for (root, dirs, files) in os.walk(path):
        # Append folder name
        total_information = total_information + "\n" + ('Folder: ' + root)

        folder_layer = root.count("\\")
        listset = ['Folder Path', root, '', '', '', folder_layer]
        folder_dict = append_dic(folder_dict, keyset, listset)

        # if it contains subFolder
        if dirs:
            # collect subfolder names
            folder_dic = 'It contains subfolder: '
            for names in dirs:
                folder_dic += names + ' , '
            folder_dic = folder_dic[:-2]
            total_information = total_information + "\n" + folder_dic

            # Collect file size
            for names in dirs:
                subfolder_path = root + "\\" + names
                size = get_dir_size(subfolder_path)
                total_information = total_information + "\n" + \
                                    "The size of folder " + names + " is " + str(round(size / 1024, 1)) + ' kb'

                # Add info to dictionary
                listset = ['Subfolder', names, round(size / 1024, 1), '', '', '']
                folder_dict = append_dic(folder_dict, keyset, listset)
        else:
            total_information = total_information + "\n" + 'It contains no subfolder: '
            listset = ['Subfolder', 'No folder', '0', '', '', '']
            folder_dict = append_dic(folder_dict, keyset, listset)

        # if it contains files
        if files:
            file_dic = 'It contains files: '
            for names in files:
                file_dic += names + ' , '
            file_dic = file_dic[:-2]
            total_information = total_information + "\n" + file_dic

            # Append file size
            for names in files:
                file_path = root + "/" + names
                total_information = total_information + "\n" + 'File size of ' \
                                    + names + ' is ' + str(round(os.path.getsize(file_path) / 1024, 1)) + ' kb'
                size_str = str(round(os.path.getsize(file_path) / 1024, 1)) + ' kb'

                # Add info to dictionary
                if round(os.path.getsize(file_path) / 1024, 1) < check_size:
                    check = 'Check'
                else:
                    check = ''

                listset = ['File', names, round(os.path.getsize(file_path) / 1024, 1), '', check, '']
                folder_dict = append_dic(folder_dict, keyset, listset)

            # Count total file number

            total_information = total_information + "\n" + "Total files in the folder is: " + str(len(files))

            listset = ['File num', '', '', len(files), '', '']
            folder_dict = append_dic(folder_dict, keyset, listset)
        else:
            total_information = total_information + "\n" + 'It contains no files: '

            listset = ['File', "No file", '', '0', '', '']
            folder_dict = append_dic(folder_dict, keyset, listset)

        # Leave a blank line between each folder
        listset = ['', "", '', '', '', '']
        folder_dict = append_dic(folder_dict, keyset, listset)

        # append separate lines
        total_information = total_information + "\n" + '--------------------------------'

    folder_info_str = total_information[1:]

    # Turn the diction to dataframe
    folder_info_df = pd.DataFrame(folder_dict)

    return folder_info_str, folder_info_df


# collect folder information.
# csv, true means result in csv
# txt, true means result in txt
def info_to_text_csv(path, csv, txt, folder_dict, keyset, check_size):
    txt_info, df_info = collect_info(path, folder_dict, keyset, check_size)
    # Write information to csv file
    if csv:
        with pd.option_context('display.max_rows', None,
                               'display.max_columns', None,
                               'display.precision', 3,
                               ):
            print(df_info)
        df_info.to_csv('folder_info.csv', index=False)

    if csv and txt:
        print("\n-----------------separation-----------------\n")

    # Write information to text file
    if txt:
        text_file = open("folder_info.txt", "w")
        text_file.write(txt_info)
        text_file.close()
        print(txt_info)


if __name__ == "__main__":
    folder_dict = {'Category': [],
                   'Names': [],
                   'Size': [],
                   'Count': [],
                   '<10kb': [],
                   'Subfolder Level': []}
    key_set = ['Category', 'Names', 'Size', 'Count', '<10kb', 'Subfolder Level', ]
    path = "demo1"
    info_to_text_csv(path=path, csv=True,
                     txt=True, folder_dict=folder_dict,
                     keyset=key_set, check_size=10)
