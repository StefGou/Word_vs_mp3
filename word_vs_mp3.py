import datetime, os, shutil

#Iterate over the files in a directory and append them in a list.
def create_list(folder):
    files = []

    for file in os.listdir(folder):
        #Append only if it's a file, not a directory
        if os.path.isdir(os.path.join(folder,file)):
            continue
        files.append(file)

    return files


#Extrat the date form the Word file's name. e.g. Category_16-12-31_Subject.docx ==> 16-12-31
def extract_date_word(word):
    return word.split("_")[1].split(".docx")[0].split(" ")[0]

#Extrat the date form the mp3 file's name. e.g. 2016-12-31 - Category.mp3 ==>  ==> 2016-12-31
def extract_date_mp3(mp3):
    return mp3.split(" - ")[0]

#Extrat the name (Category) form the Word file's folder name. e.g. D:\folder\directory\Category ==> Category
def extract_name_word(word):
    return word.split("\\")[-1]


#Verify if the Word file and the mp3 file were created less than 7 days appart.
def in_between(word,mp3):
    #Exclude "template" files.
    if not "MODÈLE" in word:
        #Convert the date to a datetime object
        word_date = datetime.datetime.strptime(extract_date_word(word), "%y-%m-%d").date()

        #Add 7 days to assign the end date.
        date_end = word_date + datetime.timedelta(days=7)

        #Convert the date to a datetime object
        mp3_date = datetime.datetime.strptime(extract_date_mp3(mp3), "%Y-%m-%d").date()

        #True if the mp3's date is within 7 days of the Word file's date
        return word_date <= mp3_date <= date_end


#Copy matching mp3's in the specified directory.
def copy_my_mp3(mp3, word, folder):
    dest_dir = folder+"\Mes sketchs"

    #If the directory does'nt exist, create it.
    if not os.path.isdir(dest_dir):
        os.makedirs(dest_dir)

    #If the mp3 is'nt already copied, copy it.
    if not os.path.isfile(dest_dir + "\\" + mp3):
        shutil.copy(folder+"\\"+mp3, dest_dir)
        print("MATCH! : {}   ===   {}".format(word,mp3))
        print("COPIE EN COURS")


#Compare two directories, looking for dates (within 7 days) and names that match.
def compare(folder_word, folder_mp3):
    list_words = create_list(folder_word)
    list_mp3s = create_list(folder_mp3)

    for word in list_words:
        for mp3 in list_mp3s:
            if in_between(word, mp3) and extract_name_word(folder_word).lower() in mp3.lower():
                copy_my_mp3(mp3, word, folder_mp3)


#Iterate over all the folders in the main Word directory
def analyse_dir(main_dir, dir_mp3):
    for folder in os.listdir(main_dir):
        #Tell the user wich directory whe are analysing
        print("Analyse du dossier: " + folder)
        #Compare the Word directory with the mp3 directory
        compare(main_dir + "\\" + folder, dir_mp3)



#Main directory containing multiple directories of Word files.
main_word = "C:\Directory\Containing\Directories"
#Main directory containing all the mp3's
main_mp3 = "C:\Directory\Containing\Mp3s"

analyse_dir(main_word,main_mp3)