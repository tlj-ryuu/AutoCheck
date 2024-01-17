import os
import win32com.client as win32

# 应对一下filepath是空的情况以及open的操作是空的情况 尽量用try语句

file_path = r'D:\TestAutoCheck'
obj = None
app = None
last_filename = '' # full path of last

class GuiSimulate():

    def __init__(self):
        self.obj = None # 当前打开的对象
        self.app = None # 当前打开的app

        # 或许下面两端 单独弄成 按钮响应事件
        self.file_path = '' # 暂时没搞定怎么获取
        self.docxs, self.excels, self.pdfs = get_filename(self.file_path)
        self.last_filename = ''

    def onButton1(self):
        pass

    def onButton2(self):
        pass

# ! 定义成按钮事件函数
def open_file(open_signal):
    """
    :param open_signal: 1:word;2:excel;3:ppt
    :param file_path: where files belongs to
    :return:
    """

    global obj
    global app
    global last_filename

    # 当obj为不空时
    if obj != None and app != None:
        close_file(last_filename)

    filename = next_file(last_filename, open_signal)
    last_filename = filename # update
    print(filename)

    if open_signal == 1:
        obj, app = open_word(filename)

    elif open_signal == 2:
        obj, app = open_excel(filename)

    elif open_signal == 3:
        obj, app = open_ppt(filename)

    else:
        pass
        obj, app = '', ''

def close_file(last_filename):
    global obj
    global app
    filename = last_filename
    ext = filename.split('.')[1]
    if ext == 'docx':
        close_word(obj, app)
    elif ext == 'xlsx' or ext == 'xls':
        close_excel(obj, app)

    elif ext == 'ppt' or ext == 'pptx':
        close_ppt(obj, app)

def get_filename(file_path):
    docxs = []
    excels = []
    ppts = []
    for root, dirs, files in os.walk(file_path):
        for file in files:
            ext = file.split('.')[1]
            if ext == 'docx':
                docxs.append(os.path.join(root, file))
            elif ext == 'xls' or ext == 'xlsx':
                excels.append(os.path.join(root, file))
            elif ext == 'ppt' or ext == 'pptx':
                ppts.append(os.path.join(root, file))
        # Just recurse one level
        break

    return docxs, excels, ppts

docxs, excels, ppts = get_filename(file_path)
print(docxs)
print(excels)
print(ppts)

def next_file(last_filename, open_signal):
    global docxs
    global excels
    global ppts
    if last_filename == '':
        if open_signal == 1:
            return docxs[0]
        elif open_signal == 2:
            return excels[0]
        elif open_signal == 3:
            return ppts[0]
    else:
        filename = os.path.split(last_filename)[1]
        ext = filename.split('.')[1]
        if ext == 'docx':
            for i in range(len(docxs)):
                if filename == os.path.split(docxs[i])[1]:
                    return docxs[(i + 1) % len(docxs)]
        elif ext == 'xls' or ext == 'xlsx':
            for i in range(len(excels)):
                if filename == os.path.split(excels[i])[1]:
                    return excels[(i + 1) % len(excels)]
        elif ext == 'ppt' or ext == 'pptx':
            for i in range(len(ppts)):
                if filename == os.path.split(ppts[i])[1]:
                    return ppts[(i + 1) % len(ppts)]

def open_word(filename):
    """
    !!! only apply to win environment and word doc of
        Microsoft Office whose extension is docx
    !!!
    """

    # Create word application objects
    word_app = win32.DispatchEx('Word.Application') # This solves process-related problems when make successive calls
    # word_app = win32.Dispatch('wps.Application')
    # print(dir(word_app))


    # Open the word window explicitly
    word_app.visible = True
    word_app.DisplayAlerts = 0

    # Open word file
    doc = word_app.Documents.Open(filename)
    print(doc)
    print('docx file has been opened')


    # 插入判断逻辑 保证判断完后关闭 要与pyqt5的信号打交道
    react = input('---------------------:')
    # 判断出结果，把word标记为红色
    if react == '1':
        word_app.ActiveWindow.View.DisplayBackgrounds = True  # 这句很重要
        red_color = 255 + (0 * 256) + (0 * 256 * 256)
        change_word_background_color(doc, red_color)
        doc.Save()
        print('docx Sucessfully Changed!')

    return doc, word_app

def close_word(doc, word_app):
    doc.Close()
    print('docx file has been closed')
    # release source
    word_app.Quit()


# 定义成按钮函数
def change_word_background_color(doc, rgb_color):
    doc.Background.Fill.ForeColor.RGB = rgb_color
    doc.Background.Fill.Visible = -1
    doc.Background.Fill.Solid()


def open_excel(filename):
    """
    !!! only apply to win environment and excel of
        Microsoft Office whose extension is xls or xlsx
    !!!
    """

    # Create excel application objects
    excel_app = win32.DispatchEx('Excel.Application') # This solves process-related problems when make successive calls

    # Open the word window explicitly
    excel_app.visible = True
    excel_app.DisplayAlerts = 0

    # Open excel file
    excel = excel_app.Workbooks.Open(filename)
    print(excel)
    print('excel file has been opened')

    # 插入判断逻辑 保证判断完后关闭 要与pyqt5的信号打交道
    react = input('---------------------:')
    # 判断出结果，把excel背景标记为红色
    if react == '1':
        red_color = 255 + (0 * 256) + (0 * 256 * 256)
        change_excel_background_color(excel, red_color)
        excel.Save()
        print('excel Sucessfully Changed!')

    return excel, excel_app


# 定义成按钮函数
def change_excel_background_color(excel, rgb_color):
    worksheet = excel.Worksheets(1)
    # Gets the scope of the entire table with filled content
    table_range = worksheet.UsedRange
    # Set the background color of the entire table to red (RGB: 255, 0, 0)
    table_range.Interior.Color = rgb_color


def close_excel(excel, excel_app):
    excel.Close()
    print('excel file has been closed')
    # release source
    excel_app.Quit()



def open_ppt(filename):
    # Create a PowerPoint application object
    ppt_app = win32.DispatchEx("PowerPoint.Application")

    # Set  PowerPoint  visible
    ppt_app.Visible = True
    ppt_app.DisplayAlerts = 0

    # Create a new presentation
    ppt = ppt_app.Presentations.Open(filename)

    print(ppt)
    print('ppt file has been opened')

    # 插入判断逻辑 保证判断完后关闭 要与pyqt5的信号打交道
    react = input('---------------------:')
    # 判断出结果，把excel背景标记为红色
    if react == '1':
        red_color = 255 + (0 * 256) + (0 * 256 * 256)
        insert_red_cross_on_first_slide(ppt, red_color)
        ppt.Save()
        print('ppt Sucessfully Changed!')

    return ppt, ppt_app

def close_ppt(ppt, ppt_app):
    ppt.Close()
    print('ppt file has been closed')
    # release source
    ppt_app.Quit()

def insert_red_cross_on_first_slide(ppt, rgb_color):

    # Get the first slide
    first_slide = ppt.Slides(1)

    # Insert the Red Cross figure
    left = 100  # Top left X coordinates
    top = 100   # Top left Y coordinates
    width = 50  # figure wide
    height = 50 # figure height

    red_cross_shape = first_slide.Shapes.AddShape(165, left, top, width, height)

    # Set the fill color of the fork graphics to red
    red_cross_shape.Fill.ForeColor.RGB = rgb_color



if __name__ == '__main__':
    for i in range(2):
        print(f'round-----------------{i}------------------')
        open_file(3)

    # word, word_app = open_word('D:\\TestAutoCheck\\Doc1.docx')
    # print(type(word))
    # print(type(word_app))
    # # open_word('D:\\TestAutoCheck\\Doc1.docx')
    # # open_word('D:\\TestAutoCheck\\zhis is test2.docx')
    # excel, excel_app = open_excel('D:\\TestAutoCheck\\Excel2.xls')
    # print(type(excel))
    # print(type(excel_app))
    # # open_word('D:\\TestAutoCheck\\Excel1.xlsx')
    # ppt, ppt_app = open_ppt('D:\\TestAutoCheck\\PPT2.ppt')
    # print(type(ppt))
    # print(type(ppt_app))