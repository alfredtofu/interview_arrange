import argparse
import xlwt
import xlrd
from datetime import datetime, timedelta

specify_final = False
interviewers = None
interviewees = None


def parse_args():
    parser = argparse.ArgumentParser('校招面试安排')
    parser.add_argument('input_file', help='输入文件（excel文件，第一行面试官，第二行终面官，第三行面试者）')
    parser.add_argument('--start_time', help='开始时间', default='9:00')
    parser.add_argument('--interval_time', help='时间间隔（分钟）', default=30)
    return parser.parse_args()


def read_data(input_file):
    global interviewees, interviewers
    input_book = xlrd.open_workbook(input_file).sheet_by_index(0)

    interviewees = [[], []]
    for i, interviewee in enumerate(input_book.col_values(0)):
        if interviewee:
            interviewees[0].append(interviewee)

    for i, interviewee in enumerate(input_book.col_values(1)):
        if interviewee:
            interviewees[1].append(interviewee)

    interviewers = []
    for i, interviewer in enumerate(input_book.col_values(2)):
        if interviewer:
            interviewers.append(interviewer)


if __name__ == '__main__':
    args = parse_args()
    start_time = datetime.strptime(args.start_time, '%H:%M')
    args.interval_time = int(args.interval_time)

    read_data(args.input_file)
    specify_final = len(interviewees[1]) != 0

    output_book = xlwt.Workbook()
    sheet1 = output_book.add_sheet('面试安排', cell_overwrite_ok=True)
    idx = 1
    sheet1.write(0, 0, '面试官')
    for i in range(len(interviewees)):
        for j in range(len(interviewees[i])):
            sheet1.write(idx, 0, interviewees[i][j])
            idx += 1

    step = 2 if len(interviewees[1]) != 0 else 3
    max_col = 0
    cur_col = [0] * len(interviewees[1])
    for i in range(len(interviewers)):
        beg_x = int(i % len(interviewees[0]))
        beg_y = int(i / len(interviewees[0])) * step + 1
        sheet1.write(beg_x % len(interviewees[0]) + 1, beg_y, interviewers[i])
        sheet1.write((beg_x + 1) % len(interviewees[0]) + 1, beg_y + 1, interviewers[i])
        if step == 2:
            beg_x_f = int(i % len(interviewees[1]))
            beg_y_f = int(i / len(interviewees[1]))
            if beg_y_f == 0:
                cur_col[beg_x_f] = int(i / len(interviewees[1])) + beg_y + 1 + 1
            else:
                cur_col[beg_x_f] = max(cur_col[beg_x_f] + 1, beg_y + 1 + 1)
            sheet1.write(beg_x_f + len(interviewees[0]) + 1,
                         cur_col[beg_x_f],
                         interviewers[i])
            max_col = cur_col[beg_x_f]
        else:
            max_col = beg_y + 2
            sheet1.write((beg_x + 2) % len(interviewees[0]) + 1, max_col, interviewers[i])
    for i in range(1, max_col):
        sheet1.write(0, i,
                     (start_time + timedelta(minutes=(i - 1) * args.interval_time)).time().strftime('%H:%M') + '-' +
                     (start_time + timedelta(minutes=i * args.interval_time)).time().strftime('%H:%M'))
    output_book.save('面试安排.xls')
