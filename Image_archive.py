import pandas as pd
import os
import shutil


class Image_archive():
    def __init__(self,path_excelfile, path_image_folder):
        self.path_excelfile=path_excelfile
        self.path_image_folder=path_image_folder
        self.data_frame=pd.read_excel(path_excelfile)
    
    
    def copy_images_by_person_to_Desktop(self,person_name):
        # create new folder on desktop
        new_folder=Image_archive.create_desktop_folder(person_name)

        # get list with images names containing person with person_name
        image_names=self.get_filenames_from_HHVL_archive(person_name)

        # get all filenames of images from path_image_folder
        file_list=Image_archive.get_all_files(self.path_image_folder)
        
        for image_name in image_names:
            try:
                file_path=Image_archive.find_image_in_file_list(image_name, file_list)
                # copy file to new folder
                shutil.copyfile(file_path, new_folder+'\\'+image_name+'.jpg')
            except: print('Some images could not be found')

    
    def get_filenames_from_HHVL_archive(self, person_name):
        '''read excel file from HHVL archive and get all filenames of images that contain the persons with person_name
        The excel file contains a column Bildname and a column Personen
        Input: person_name, string: name of person to be searched for
        Output: image_names, list of strings: filenames of images containing person_name
        '''
        image_names=[]
        for index, persons in enumerate(self.data_frame['Personen']): # go through column
            if type(persons)==str:
                if person_name in persons:  
                       image_names.append(self.data_frame['Bildname'][index])
        return image_names

    
    
    # general functions-----------------------------------------------------------------------------
    def get_all_files(path):
        '''Get the list of all files in directory tree at given path
        Input: path, string: path to folder containing files and subfolders
        Output: ListOfFiles: List containing all the files
        ''' 
        listOfFiles = list()
        for (dirpath, dirnames, filenames) in os.walk(path):
            listOfFiles += [os.path.join(dirpath, file) for file in filenames]
        return listOfFiles
    
    def create_desktop_folder(folder_name):
        '''Create folder on Desktop if it does not exist.
        Input: folder_name, string: name of folder
        Output: path to new folder
        '''
        folder_path = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop')+'\\'+folder_name
        if not os.path.exists(folder_path):
            os.makedirs(folder_path)
        return folder_path

    def find_image_in_file_list(image_name, file_list):
        '''Find image with image_name in file_list
           Input: image_name, string
           Output: path to file
        '''
        for filename in file_list:
            if '\\' + image_name +  '.' in filename: # guarantee that exactly name and not only substring
                return filename
                

