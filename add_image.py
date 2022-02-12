'''
1. Place the files in the directory "raw_files" and find the output files
   in "output" directory.
2. Place the images in "esig" directory
'''

from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from openpyxl.worksheet.worksheet import Worksheet

def execute(filename: str) -> None:
    FCU_SHEET_LAST_INDEX = 14

    wb = load_workbook(f'raw_files/{filename}.xlsx')

    def insertimage(image_loc: str, worksheet: Worksheet, cell: str) -> None:
        image = Image(image_loc)
        worksheet.add_image(image, cell)

    image1 = 'esig/jec.jpg'
    image2 = 'esig/dmta.png'
    for i in range(FCU_SHEET_LAST_INDEX, len(wb.worksheets)):
        insertimage(image1, wb.worksheets[i], 'D60')
        insertimage(image2, wb.worksheets[i], 'B60')

    wb.save(f'output/{filename}_with_esig.xlsx')
    print(f'E-sig added for {filename}')

if __name__ == '__main__':
    filenames = ['tower', 'podium', 'basement']

    for filename in filenames:
        execute(filename)
