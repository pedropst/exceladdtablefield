from msilib.schema import Class
import openpyxl
import os
from pathlib import Path
import tkinter

BASE_DIR = Path(__file__).resolve().parent

def get_category_dictionary(path):
    cat_dic = {}
    if os.path.exists(os.path.join(path, 'category_dictionary.xlsx')):
        category_book = openpyxl.load_workbook(os.path.join(path, 'category_dictionary.xlsx'))
        book = category_book.active

        for rows in book.iter_rows(min_col=1, max_col=1, min_row=1, max_row=book.max_row):
            for cell in rows:
                cat_dic[cell.value] = book.cell(cell.row, 2).value

        cat_dic.pop('MERCHANT')
        return cat_dic
    else:
        book = openpyxl.Workbook()
        sheet = book.active
        sheet.title = "Category Dictionary"
        sheet.cell(1, 1, 'MERCHANT')
        sheet.cell(1, 2, 'CATEGORY')
        book.save(os.path.join(path, 'category_dictionary.xlsx'))
        get_category_dictionary(path, cat_dic)

def add_to_category_dictionary(path, list_to_add):
    category_book = openpyxl.load_workbook(os.path.join(path, 'category_dictionary.xlsx'))
    book = category_book.active

    for k in list_to_add.keys():
        book.cell(book.max_row + 1, 1, k)
        book.cell(book.max_row, 2, list_to_add[k])

    category_book.save(os.path.join(path, 'category_dictionary.xlsx'))

class Categories:
    def __init__(self):
        self.input_button = None
        self.choosed = ''
        self.run()


    def run(self):
        path = 'samplefile.xlsx'
        books = openpyxl.load_workbook(path)
        target_column = None
        category_column = None
        merchant_column = None

        merchant_content = {}

        for book in books:
            # if book.namespace == 'Report':
            for c in book.iter_cols(min_col=0, max_col=len(book.column_dimensions), min_row=1, max_row=1):
                for cell in c:
                    if str(cell.value) == 'Description':
                        target_column = cell
                        for r in book.iter_rows(min_row=2, max_row=book.max_row, min_col=cell.column, max_col=cell.column):
                            for cell1 in r:
                                content = cell1.value
                                if content != None:
                                    if content.__contains__(',USD '):
                                        merchant_content[cell1] = content.split(',USD ')[1][:-3]
                                    elif content.__contains__(',AED '):
                                        merchant_content[cell1] = content.split(',AED ')[1][:-3]
                                else:
                                    break

            i = cell.coordinate
            f = book.dimensions.split(':')[1]

            select_merged_cells = []
            for merged_cell in book.merged_cells.ranges:
                if merged_cell.min_col >= target_column.column:
                    select_merged_cells.append(merged_cell)

            for mc in select_merged_cells:
                mc.shift(2,0)

            book.insert_cols(target_column.column, 2)

            for col in book.column_dimensions:
                book.column_dimensions[col].width = 14.28

            category_column = target_column.column - 2
            book.cell(1, category_column, 'Category')


            merchant_column = target_column.column - 1
            book.cell(1, merchant_column, 'Merchant')

            # book.column_dimensions[category_column].bestFit = True
            # book.column_dimensions[merchant_column].bestFit = True

            for row in book.iter_rows(min_row=2, max_row=book.max_row, min_col=category_column, max_col=category_column):
                for cell in row:
                    cell.value = "teste"

            category_dictionary = {'MERCHANT' : 'CATEGORY'}
            old_data = {}

            def call(entry):
                self.choosed = entry.get()
                self.input_button.destroy()
                self.input_button = None

            new_data = {}
            for k in merchant_content.keys():
                old_data = get_category_dictionary(BASE_DIR)
                book.cell(k.row, merchant_column, merchant_content[k])
                if not old_data.keys().__contains__(merchant_content[k]):
                    # OPEN A SIMPLE INPUT FIELD
                    self.input_button = tkinter.Tk()
                    self.input_button.geometry('500x200')
                    self.input_button.title('Add new category or assign')
                    l = tkinter.Label(self.input_button, text=merchant_content[k])
                    l.pack()
                    e = tkinter.Entry(self.input_button, width=20)
                    e.pack()
                    b = tkinter.Button(self.input_button, text='OK', command=lambda: call(e))
                    b.pack()
                    self.input_button.mainloop()

                    while self.input_button != None:
                        continue
                    
                    new_data[merchant_content[k]] = self.choosed
                    self.choosed = ''
                else:
                    new_data[merchant_content[k]] = old_data[merchant_content[k]]
            
            add_to_category_dictionary(BASE_DIR, new_data)


            print('ffasdas')

                    
            # book.move_range(f"{target_column.coordinate}:{book.dimensions.split(':')[1]}", rows=0, cols=20, translate=True)

            # book.insert_cols(target_column)
            # category_column = target_column
            # book.cell(1, category_column, 'Category')

            # book.insert_cols(target_column - 1)
            # merchant_column = target_column - 1
            # book.cell(1, merchant_column, 'Merchant')
            
            # target_column += 1

        books.save('result.xlsx')


c = Categories()