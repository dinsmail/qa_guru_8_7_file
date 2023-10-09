import os
from zipfile import ZipFile
from openpyxl import load_workbook

zipFilesDir = './resources'
zipPath = './tmp/testzip.zip'


def create_zip():
    with ZipFile(zipPath, 'w') as testZip:
        for filename in os.listdir(zipFilesDir):
            filename_with_dir = os.path.join(zipFilesDir, filename)
            filename_full = os.path.abspath(filename_with_dir)
            testZip.write(filename_full, filename)


def delete_zip():
    os.remove(zipPath)


def test_check_zip_file_is_created():
    create_zip()
    assert os.path.isfile(zipPath)


def test_number_of_files_is_same():
    create_zip()
    with ZipFile(zipPath, mode='r') as testZip:
        assert len(testZip.namelist()) == len(os.listdir(zipFilesDir))


def test_files_are_same():
    create_zip()
    with ZipFile(zipPath, 'r') as testZip:
        # check all files from directory exists in zip
        for filename in os.listdir(zipFilesDir):
            assert filename in testZip.namelist()
        # check all files in zip have same length as file in directory
        for zip_item in testZip.filelist:
            filename = os.path.join(os.path.join(zipFilesDir, zip_item.filename))
            zip_size = zip_item.file_size
            file_size = os.path.getsize(filename)
            assert zip_size == file_size


def test_check_zip_is_removed():
    delete_zip()
    assert not os.path.isfile(zipPath)
