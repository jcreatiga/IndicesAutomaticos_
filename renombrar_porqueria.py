import os

SIMILARITY_THRESHOLD = 1
destination_file = "00IndiceExpedienteElectronico.xlsx"

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'


def similarity(word1, word2):
    """
    Calculates the similarity between two words using the Levenshtein distance algorithm.
    Returns a value between 0 and 1, where 0 means the words are completely different and 1 means they are identical.
    """
    if len(word1) < len(word2):
        return similarity(word2, word1)

    if len(word2) == 0:
        return 1.0

    previous_row = range(len(word2) + 1)
    for i, c1 in enumerate(word1):
        current_row = [i + 1]
        for j, c2 in enumerate(word2):
            insertions = previous_row[j + 1] + 1
            deletions = current_row[j] + 1
            substitutions = previous_row[j] + (c1 != c2)
            current_row.append(min(insertions, deletions, substitutions))
        previous_row = current_row

    return 1.0 - (previous_row[-1] / max(len(word1), len(word2)))

def check_nonunified_files():
    conflict_files = []
    
    for dirpath, _, files in os.walk("."):
        # check if there is a file with .xlsx extension
        for file in files:
            if file.endswith(".xlsx") and file != destination_file and similarity(file.lower(), destination_file.lower()) >= 0.3:
                conflict_files.append(os.path.join(dirpath, file))
                try:
                    os.rename(os.path.join(dirpath, file), os.path.join(dirpath, destination_file))
                except Exception as e:
                    # delete the file
                    try:
                        os.remove(os.path.join(dirpath, file))
                    except Exception as e:
                        print(bcolors.WARNING + "ERROR: No se pudo eliminar el archivo: " + os.path.join(dirpath, file))

def delete_files():
    conflict_files = []
    
    for dirpath, _, files in os.walk("."):
        # check if there is a file with .xlsx extension
        for file in files:
            if file == destination_file:
                try:
                    os.remove(os.path.join(dirpath, file))
                    conflict_files.append(os.path.join(dirpath, file))
                    print(bcolors.OKGREEN + "Archivo eliminado: " + os.path.join(dirpath, file))
                except Exception as e:
                    print(bcolors.WARNING + "ERROR: No se pudo eliminar el archivo: " + os.path.join(dirpath, file))
                
    print(bcolors.OKGREEN + "Se eliminaron " + str(len(conflict_files)) + " archivos.")

delete_files()