import csv
import hashlib
import json
import os
import subprocess
import matplotlib.pyplot as plt  # Importing matplotlib
import pandas as pd
import xlrd2
from olefile import olefile
from olefile.olefile import *
from oletools import oleobj, olevba


def clean_redundant_files(dir_path: str):
    """
    Clean files that are created due to the oleobj analysis run.
    :param dir_path:
    :return:
    """
    # List of valid lowercase extensions
    valid_extensions = {'.ppt', '.xls', '.doc'}

    # Ensure the directory path is valid
    if not os.path.exists(dir_path):
        raise ValueError("Directory path does not exist.")

    # Iterate through the files in the directory
    for filename in os.listdir(dir_path):
        file_path = os.path.join(dir_path, filename)

        # Check if it's a file and not a directory
        if os.path.isfile(file_path):
            # Check if the file extension is not in the valid extensions set
            if not filename.lower().endswith(tuple(valid_extensions)):
                try:
                    # Delete the file if it doesn't meet the condition
                    os.remove(file_path)
                    print(f"Deleted: {file_path}")
                except Exception as e:
                    print(f"Error deleting {file_path}: {e}")


def get_corrupted_files(ppt_dir):
    """
    Return a tuple of two lists of corrupted ppt files: one of files that doesn't contain powerpoint document stream,
    and the second list contain files that has malformed embedded objects.
    :param ppt_dir: dir contains the original ppt files.
    :return:
    """
    cor_files_sha = []
    bad_powdoc_files_sha = []

    for file in os.listdir(ppt_dir):
        file_path = os.path.join(ppt_dir, file)
        file_type = file_path.split('.')[-1].lower()
        file_orig_sha = calc_sha256(file_path)

        if file_type != 'ppt':
            continue

        print("Processing file:", file_path)
        try:
            ole = olefile.OleFileIO(filename=file_path)
        except:
            print("Invalid file!")
            continue

        if not ole.exists('PowerPoint Document'):
            print("Corrupted file")
            cor_files_sha.append(file_orig_sha)
        else:
            output_obj_str = subprocess.run(f'oleobj {file_path}', shell=True, capture_output=True, text=True).stdout

            if "Error reading data from PowerPoint Document stream or interpreting it as OLE object" in output_obj_str:
                print("error powerpoint document")
                bad_powdoc_files_sha.append(file_orig_sha)

    clean_redundant_files(ppt_dir)

    return cor_files_sha, bad_powdoc_files_sha


def calc_sha256(filepath):
    print(filepath)
    try:
        with open(filepath, "rb") as f:
            bytes = f.read()  # read entire file as bytes
            readable_hash = hashlib.sha256(bytes).hexdigest()
            return readable_hash
    except:
        pass


def get_objects_amount(filename: str) -> int:
    """
    Function counts the amount of objects in file.
    The amount is considered by: xlm macros, VBA modules, and 'ole10native' streams.
    :param filename: file's path
    :return: the amount of objects in file.
    """
    modules_cnt = 0
    macro_sheets_cnt = 0
    ole_native_cnt = 0

    for ole in oleobj.find_ole(filename, None):
        if ole is None:
            continue

        ole.fp.seek(0)
        ole_data = ole.fp.read()

        try:
            vba_parser = olevba.VBA_Parser(filename, data=ole_data)
            vba_parser.detect_macros()
            vba_parser.extract_all_macros()
        except:
            continue

        # Count the modules amount inside the vba project of the current ole file
        for subfilename, stream_path, vba_filename, vba_code in vba_parser.modules:
            if subfilename == "xlm_macro":
                continue
            modules_cnt += 1

        for line in vba_parser.xlm_macros:
            if "Sheet Information - Excel 4.0 macro sheet" in line:
                macro_sheets_cnt += 1

        # Search for \x01Ole10Native stream
        for direntry in ole.direntries:
            if direntry is None:
                continue
            if direntry.name.lower() == "\x01ole10native":
                ole_native_cnt += 1

    obj_cnt = modules_cnt + macro_sheets_cnt + ole_native_cnt
    print(f"modules: {modules_cnt}, xlm macro sheets: {macro_sheets_cnt}, ole10native objects: {ole_native_cnt}")
    print("Total objects:", obj_cnt)

    return obj_cnt


def get_pages_amount(filename: str, file_type: str):
    """
    Get the amount of pages from olefile as following:
        XLS -> worksheets number
        PPT -> slides property
        DOC -> pages property
    :param filename:
    :return: the file's pages.
    """
    try:
        ole = OleFileIO(filename=filename)
        meta = ole.get_metadata()
    except:
        return -1

    nm_pages = 0
    pages = meta.num_pages
    slides = meta.slides

    try:  # Might be failed
        # XLS FILE
        if file_type == 'xls':
            xls_file_book = xlrd2.open_workbook(file_path)
            print("The number of worksheets is {0}".format(xls_file_book.nsheets))
            nm_pages = xls_file_book.nsheets

        # PPT FILE
        if file_type == 'ppt' and slides >= 0:
            print("The number of slides is {0}".format(slides))
            nm_pages = 1 if slides == 0 else slides

        # DOC FILE
        if file_type == 'doc' and pages >= 0:
            print("The number of pages is {0}".format(pages))
            nm_pages = pages
    except:
        return -1

    return nm_pages


def set_plt_settings():
    # Set a gray background
    plt.gca().set_facecolor('lightgray')

    # Add white grid lines to the plot
    plt.grid(True, color='white', linestyle='-', linewidth=0.5)

    # Custom legend labels
    legend_labels = {'PowerPoint': 'orange', 'Excel': 'green', 'Word': 'blue'}

    handles = [plt.Rectangle((0, 0), 1, 1, color=color, label=label) for label, color in legend_labels.items()]
    plt.legend(handles=handles, loc='lower center', bbox_to_anchor=(0.5, -0.2), ncol=len(legend_labels), borderaxespad=-0.2)


if __name__ == '__main__':
    javacdr_metadata_path = '/home/amir/Downloads/Result-CDR/metadata'
    pycdr_metadata_path = '/home/amir/Downloads/Result-CDR/pycdr_metadata'
    cdr_dir = '/home/amir/Downloads/Result-CDR/CDR'
    cdr_orig_dir = '/home/amir/Downloads/Result-CDR/CDR-org'
    ppt_dir = '../Result/OLE/PPT'

    files_dict = {
        "doc": [[], [], []],
        "ppt": [[], [], []],
        "xls": [[], [], []]
    }

    data = []

    # Corrupted files sha list
    cor_files_sha, bad_powdoc_files_sha = get_corrupted_files(ppt_dir)
    cor_files_sha_lst = cor_files_sha + bad_powdoc_files_sha

    for file in os.listdir(cdr_dir):
        print("processing file:", file)

        file_path = os.path.join(cdr_orig_dir, file)
        file_type = file.split('.')[-1].lower()
        file_orig_sha = calc_sha256(file_path)
        print("sha:", file_orig_sha)

        # Ignoring corrupted files in malware dataset
        if file_orig_sha in cor_files_sha_lst:
            continue

        try:
            # Java cdr file's metadata
            jcdr_meta_file_path = f'{os.path.join(javacdr_metadata_path, file)}.json'
            with open(jcdr_meta_file_path, 'r') as f:
                jcdr_file_meta = json.load(f)

            # Python cdr file's metadata
            pycdr_meta_file_path = f'{os.path.join(pycdr_metadata_path, file)}.json'
            with open(pycdr_meta_file_path, 'r') as f:
                pycdr_file_meta = json.load(f)
        except:
            continue

        jcdr_analysis_duration = jcdr_file_meta['Analysis_duration']
        pycdr_analysis_duration = pycdr_file_meta['analysis_duration']
        pycdr_analysis_duration = int(pycdr_analysis_duration * 1000)  # Convert to ms

        total_analysis_duration = jcdr_analysis_duration + pycdr_analysis_duration

        if total_analysis_duration > 4000:
            continue

        file_list = files_dict[file_type]

        file_list[0].append(total_analysis_duration)

        nm_pages = get_pages_amount(file_path, file_type)
        nm_objects = get_objects_amount(file_path)

        print("Pages:", nm_pages)
        print("Objects:", nm_objects)

        file_list[1].append(nm_pages)
        file_list[2].append(nm_objects)

        # Appending a new row of file's data to an array that will be written as a csv file
        data.append({'sha256_orig': file_orig_sha, 'time': total_analysis_duration, 'pages': nm_pages, 'objects': nm_objects})

    # Save csv file of the data collected
    df = pd.DataFrame(data)
    df.to_csv('analysis_duration_malware.csv', index=False)

    # Create plot graph
    plt.scatter(files_dict['doc'][0], files_dict['doc'][1], c='blue')
    plt.scatter(files_dict['ppt'][0], files_dict['ppt'][1], c='orange')
    plt.scatter(files_dict['xls'][0], files_dict['xls'][1], c='green')

    plt.xlabel('Total Analysis Duration (ms)')
    plt.ylabel('Number of Pages')

    set_plt_settings()
    plt.savefig('pages_vs_analysis_duration.png', bbox_inches='tight', dpi=300)  # Save the figure with tight bounding box

    plt.show()

    # Create second plot graph of objects vs analysis duration
    plt.scatter(files_dict['doc'][0], files_dict['doc'][2], c='blue')
    plt.scatter(files_dict['ppt'][0], files_dict['ppt'][2], c='orange')
    plt.scatter(files_dict['xls'][0], files_dict['xls'][2], c='green')

    # Set labels for the axes
    plt.xlabel('Total Analysis Duration (ms)')
    plt.ylabel('Objects amount')

    set_plt_settings()
    plt.savefig('objects_vs_analysis_duration.png', bbox_inches='tight', dpi=300)  # Save the figure with tight bounding box
    plt.show()
