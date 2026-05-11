import shutil

def copy_directory(RennFolderPath):
    listErrors = []
    source_dir = './BesetzungenJson'
    destination_dir = RennFolderPath

    # Allows overwriting if the destination directory already exists
    shutil.copytree(source_dir, destination_dir, dirs_exist_ok=True)
    return listErrors