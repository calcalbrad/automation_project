import shutil

def copy_master_doc(original_file, new_file):
    try:
        shutil.copy(original_file, new_file)
        print(f"File copied and renamed to: {new_file}") # remove this
    except FileNotFoundError:
        print(f"The file {original_file} was not found.")
    except Exception as e:
        print(f"An error occured: {e}")