from tkinter import *
from tkinter import ttk
import pandas as pd
from matplotlib import pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

plt.rcParams['font.family'] = 'Malgun Gothic'

window = Tk()
window.geometry("1600x700+100+100")

data2 = pd.read_excel('C:/data/실습데이터/kSE수정종가_cut.xlsx', index_col=0)

path_df = ('C:/data/실습데이터/데이터프레임/')
path_stock = ('C:/data/실습데이터/주가/')

def check1():
    data = pd.read_excel(path_df + 'data_high_roe.xlsx')
    global treeview
    treeview.delete(*treeview.get_children())

    for widget in frame3.winfo_children():
        widget.destroy()
    for widget in frame4.winfo_children():
        widget.destroy()

    treeview['column'] = list(data.columns)
    treeview.column('#0', anchor='center', width=1, stretch=False)
    treeview.column('#1', anchor='center', width=100, stretch=False)
    treeview.column('#2', anchor='center', width=100, stretch=False)
    treeview.column('#9', anchor='center', width=100, stretch=False)
    treeview['show'] = 'tree headings'

    for col in treeview['column']:
        treeview.heading(col, text=col)

    data_rows = data.to_numpy().tolist()
    for row in data_rows:
        treeview.insert("", "end", values=row)

    data_high_roe_name = pd.read_excel(path_stock + "data_roe_stock.xlsx").iloc[:, 1:]
    data_high_roe_name = data_high_roe_name.columns.tolist()

    plt.clf()
    for i in range(10):
        figure = plt.Figure(figsize=(3, 2), dpi=80)
        ax = figure.add_subplot(1, 1, 1)
        if i < 5:
            line = FigureCanvasTkAgg(figure, frame3)
            line.get_tk_widget().pack(side='left')
        else:
            line = FigureCanvasTkAgg(figure, frame4)
            line.get_tk_widget().pack(side='left')
        data2.plot.line(use_index=True, y=[data_high_roe_name[i]], ax=ax)


def check2():
    data = pd.read_excel(path_df + 'data_high_roa.xlsx')
    global treeview
    treeview.delete(*treeview.get_children())

    for widget in frame3.winfo_children():
        widget.destroy()
    for widget in frame4.winfo_children():
        widget.destroy()

    treeview['column'] = list(data.columns)
    treeview.column('#0', anchor='center', width=1, stretch=False)
    treeview.column('#1', anchor='center', width=100, stretch=False)
    treeview.column('#2', anchor='center', width=100, stretch=False)
    treeview.column('#9', anchor='center', width=100, stretch=False)
    treeview['show'] = 'tree headings'

    for col in treeview['column']:
        treeview.heading(col, text=col)

    data_rows = data.to_numpy().tolist()
    for row in data_rows:
        treeview.insert("", "end", values=row)

    data_high_roa_name = pd.read_excel(path_stock + "data_roa_stock.xlsx").iloc[:, 1:]
    data_high_roa_name = data_high_roa_name.columns.tolist()

    plt.clf()
    for i in range(10):
        figure = plt.Figure(figsize=(3, 2), dpi=80)
        ax = figure.add_subplot(1, 1, 1)
        if i < 5:
            line = FigureCanvasTkAgg(figure, frame3)
            line.get_tk_widget().pack(side='left')
        else:
            line = FigureCanvasTkAgg(figure, frame4)
            line.get_tk_widget().pack(side='left')
        data2.plot.line(use_index=True, y=[data_high_roa_name[i]], ax=ax)


def check3():
    data = pd.read_excel(path_df + 'data_high_debt.xlsx')
    global treeview
    treeview.delete(*treeview.get_children())

    for widget in frame3.winfo_children():
        widget.destroy()
    for widget in frame4.winfo_children():
        widget.destroy()

    treeview['column'] = list(data.columns)
    treeview.column('#0', anchor='center', width=1, stretch=False)
    treeview.column('#1', anchor='center', width=100, stretch=False)
    treeview.column('#2', anchor='center', width=100, stretch=False)
    treeview.column('#9', anchor='center', width=100, stretch=False)
    treeview['show'] = 'tree headings'

    for col in treeview['column']:
        treeview.heading(col, text=col)

    data_rows = data.to_numpy().tolist()
    for row in data_rows:
        treeview.insert("", "end", values=row)

    data_high_debt_name = pd.read_excel(path_stock + "data_debt_stock.xlsx").iloc[:, 1:]
    data_high_debt_name = data_high_debt_name.columns.tolist()

    plt.clf()
    for i in range(10):
        figure = plt.Figure(figsize=(3, 2), dpi=80)
        ax = figure.add_subplot(1, 1, 1)
        if i < 5:
            line = FigureCanvasTkAgg(figure, frame3)
            line.get_tk_widget().pack(side='left')
        else:
            line = FigureCanvasTkAgg(figure, frame4)
            line.get_tk_widget().pack(side='left')
        data2.plot.line(use_index=True, y=[data_high_debt_name[i]], ax=ax)


def check4():
    data = pd.read_excel(path_df + 'data_high_capital.xlsx')
    global treeview
    treeview.delete(*treeview.get_children())

    for widget in frame3.winfo_children():
        widget.destroy()
    for widget in frame4.winfo_children():
        widget.destroy()

    treeview['column'] = list(data.columns)
    treeview.column('#0', anchor='center', width=1, stretch=False)
    treeview.column('#1', anchor='center', width=100, stretch=False)
    treeview.column('#2', anchor='center', width=100, stretch=False)
    treeview.column('#9', anchor='center', width=100, stretch=False)
    treeview['show'] = 'tree headings'

    for col in treeview['column']:
        treeview.heading(col, text=col)

    data_rows = data.to_numpy().tolist()
    for row in data_rows:
        treeview.insert("", "end", values=row)

    data_high_capital_name = pd.read_excel(path_stock + "data_capital_stock.xlsx").iloc[:, 1:]
    data_high_capital_name = data_high_capital_name.columns.tolist()

    plt.clf()
    for i in range(10):
        figure = plt.Figure(figsize=(3, 2), dpi=80)
        ax = figure.add_subplot(1, 1, 1)
        if i < 5:
            line = FigureCanvasTkAgg(figure, frame3)
            line.get_tk_widget().pack(side='left')
        else:
            line = FigureCanvasTkAgg(figure, frame4)
            line.get_tk_widget().pack(side='left')
        data2.plot.line(use_index=True, y=[data_high_capital_name[i]], ax=ax)


def check5():
    data = pd.read_excel(path_df + 'data_high_sales.xlsx')
    global treeview
    treeview.delete(*treeview.get_children())

    for widget in frame3.winfo_children():
        widget.destroy()
    for widget in frame4.winfo_children():
        widget.destroy()

    treeview['column'] = list(data.columns)
    treeview.column('#0', anchor='center', width=1, stretch=False)
    treeview.column('#1', anchor='center', width=100, stretch=False)
    treeview.column('#2', anchor='center', width=100, stretch=False)
    treeview.column('#9', anchor='center', width=100, stretch=False)
    treeview['show'] = 'tree headings'

    for col in treeview['column']:
        treeview.heading(col, text=col)

    data_rows = data.to_numpy().tolist()
    for row in data_rows:
        treeview.insert("", "end", values=row)

    data_high_sales_name = pd.read_excel(path_stock + "data_sales_stock.xlsx").iloc[:, 1:]
    data_high_sales_name = data_high_sales_name.columns.tolist()

    plt.clf()
    for i in range(10):
        figure = plt.Figure(figsize=(3, 2), dpi=80)
        ax = figure.add_subplot(1, 1, 1)
        if i < 5:
            line = FigureCanvasTkAgg(figure, frame3)
            line.get_tk_widget().pack(side='left')
        else:
            line = FigureCanvasTkAgg(figure, frame4)
            line.get_tk_widget().pack(side='left')
        data2.plot.line(use_index=True, y=[data_high_sales_name[i]], ax=ax)


def check6():
    data = pd.read_excel(path_df + 'data_high_op.xlsx')
    global treeview
    treeview.delete(*treeview.get_children())

    for widget in frame3.winfo_children():
        widget.destroy()
    for widget in frame4.winfo_children():
        widget.destroy()

    treeview['column'] = list(data.columns)
    treeview.column('#0', anchor='center', width=1, stretch=False)
    treeview.column('#1', anchor='center', width=100, stretch=False)
    treeview.column('#2', anchor='center', width=100, stretch=False)
    treeview.column('#9', anchor='center', width=100, stretch=False)
    treeview['show'] = 'tree headings'

    for col in treeview['column']:
        treeview.heading(col, text=col)

    data_rows = data.to_numpy().tolist()
    for row in data_rows:
        treeview.insert("", "end", values=row)

    data_high_op_name = pd.read_excel(path_stock + "data_op_stock.xlsx").iloc[:, 1:]
    data_high_op_name = data_high_op_name.columns.tolist()

    plt.clf()
    for i in range(10):
        figure = plt.Figure(figsize=(3, 2), dpi=80)
        ax = figure.add_subplot(1, 1, 1)
        if i < 5:
            line = FigureCanvasTkAgg(figure, frame3)
            line.get_tk_widget().pack(side='left')
        else:
            line = FigureCanvasTkAgg(figure, frame4)
            line.get_tk_widget().pack(side='left')
        data2.plot.line(use_index=True, y=[data_high_op_name[i]], ax=ax)


def check7():
    data = pd.read_excel(path_df + 'data_high_all.xlsx')
    global treeview
    treeview.delete(*treeview.get_children())

    for widget in frame3.winfo_children():
        widget.destroy()
    for widget in frame4.winfo_children():
        widget.destroy()

    treeview['column'] = list(data.columns)
    treeview.column('#0', anchor='center', width=1, stretch=False)
    treeview.column('#1', anchor='center', width=100, stretch=False)
    treeview.column('#2', anchor='center', width=100, stretch=False)
    treeview.column('#9', anchor='center', width=100, stretch=False)
    treeview['show'] = 'tree headings'

    for col in treeview['column']:
        treeview.heading(col, text=col)

    data_rows = data.to_numpy().tolist()
    for row in data_rows:
        treeview.insert("", "end", values=row)

    data_high_all_name = pd.read_excel(path_stock + "data_all_stock.xlsx").iloc[:, 1:]
    data_high_all_name = data_high_all_name.columns.tolist()

    plt.clf()
    for i in range(10):
        figure = plt.Figure(figsize=(3, 2), dpi=80)
        ax = figure.add_subplot(1, 1, 1)
        if i < 5:
            line = FigureCanvasTkAgg(figure, frame3)
            line.get_tk_widget().pack(side='left')
        else:
            line = FigureCanvasTkAgg(figure, frame4)
            line.get_tk_widget().pack(side='left')
        data2.plot.line(use_index=True, y=[data_high_all_name[i]], ax=ax)

RadioVar_1 = IntVar(value=1)


frame1 = Frame(window)
frame1.pack(side='top', fill='both')

label1 = Label(frame1, text='정렬 기준:')
label1.pack(side='left')

radio1 = Radiobutton(frame1, text="roe", value=1, variable=RadioVar_1, command=check1)
radio1.pack(side='left')
radio2 = Radiobutton(frame1, text="roa", value=2, variable=RadioVar_1, command=check2)
radio2.pack(side='left')
radio3 = Radiobutton(frame1, text="부채비율", value=3, variable=RadioVar_1, command=check3)
radio3.pack(side='left')
radio4 = Radiobutton(frame1, text="자기자본비율", value=4, variable=RadioVar_1, command=check4)
radio4.pack(side='left')
radio5 = Radiobutton(frame1, text="매출액증가율", value=5, variable=RadioVar_1, command=check5)
radio5.pack(side='left')
radio6 = Radiobutton(frame1, text="영업이익증가율", value=6, variable=RadioVar_1, command=check6)
radio6.pack(side='left')
radio7 = Radiobutton(frame1, text="총점", value=7, variable=RadioVar_1, command=check7)
radio7.pack(side='left')

frame2 = Frame(window)
frame2.pack(side='top')
treeview = ttk.Treeview(frame2)
treeview.pack()

frame3 = Frame(window)
frame3.pack(side='top')
frame4 = Frame(window)
frame4.pack(side='top')

data = pd.read_excel('C:/data/실습데이터/데이터프레임/data_high_roe.xlsx')
treeview.delete(*treeview.get_children())

data_high_roe_name = pd.read_excel(
        "C:/data/실습데이터/주가/data_roe_stock.xlsx").iloc[:, 1:]
data_high_roe_name = data_high_roe_name.columns.tolist()

treeview['column'] = list(data.columns)
treeview.column('#0', anchor='center', width=1, stretch=False)
treeview.column('#1', anchor='center', width=100, stretch=False)
treeview.column('#2', anchor='center', width=100, stretch=False)
treeview.column('#9', anchor='center', width=100, stretch=False)
treeview['show'] = 'tree headings'

for col in treeview['column']:
    treeview.heading(col, text=col)

data_rows = data.to_numpy().tolist()
for row in data_rows:
    treeview.insert("", "end", values=row)

plt.clf()
for i in range(10):
    figure = plt.Figure(figsize=(3, 2), dpi=80)
    ax = figure.add_subplot(1, 1, 1)
    if i < 5:
        line = FigureCanvasTkAgg(figure, frame3)
        line.get_tk_widget().pack(side='left')
    else:
        line = FigureCanvasTkAgg(figure, frame4)
        line.get_tk_widget().pack(side='left')
    data2.plot.line(use_index=True, y=[data_high_roe_name[i]], ax=ax)



# for i in range(len(data_high_roe_name)):
#     data2.plot.line(use_index=True, y=[data_high_roe_name[i]])
# for i in range(len(data_high_roa_name)):
#     data2.plot.line(use_index=True, y=[data_high_roa_name[i]])
# for i in range(len(data_high_debt_name)):
#     data2.plot.line(use_index=True, y=[data_high_debt_name[i]])
# for i in range(len(data_high_capital_name)):
#     data2.plot.line(use_index=True, y=[data_high_capital_name[i]])
# for i in range(len(data_high_sales_name)):
#     data2.plot.line(use_index=True, y=[data_high_sales_name[i]])
# for i in range(len(data_high_op_name)):
#     data2.plot.line(use_index=True, y=[data_high_op_name[i]])



window.mainloop()
