from zipfile import ZipFile

def extraction(detailed):
    # specifying the zip file name
    file_name = detailed

    # opening the zip file in READ mode
    with ZipFile(file_name, 'r') as zip:

        # extracting all the files
        print('Extracting all the files now...')
        zip.extractall('extraction')



if __name__== "__main__":
    extraction()
