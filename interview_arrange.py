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


class Person(object):
    def __init__(self, name, time_idx=0, used_lst=None, idx=-1):
        self.name = name
        self.item_idx = idx
        self.time_idx = time_idx
        self.used_lst = {} if used_lst is None else used_lst

    def __str__(self):
        return "name: %s, time_idx: %d, used_lst：%s" % (self.name, self.time_idx, str(self.used_lst))

    def __lt__(self, other):
        if len(self.used_lst) == 2 and len(self.used_lst) == len(other.used_lst) and specify_final:
            return self.item_idx < other.item_idx
        if len(interviewers) < len(interviewees[0]):
            return self.item_idx < other.item_idx
        if self.time_idx == other.time_idx:
            # 如果时间一样，看目前时间是不是刚开始
            # 如果刚开始，谁先来谁放前面
            # 如果不是刚开始，谁先来谁放后面
            return self.item_idx < other.item_idx if self.time_idx == 0 else self.item_idx > other.item_idx
        return self.time_idx > other.time_idx

    __repr__ = __str__


def read_data(input_file):
    global interviewees, interviewers
    input_book = xlrd.open_workbook(input_file).sheet_by_index(0)

    interviewees = [[], []]
    for i, interviewee in enumerate(input_book.col_values(0)):
        if interviewee:
            interviewees[0].append(Person(interviewee))

    for i, interviewee in enumerate(input_book.col_values(1)):
        if interviewee:
            interviewees[1].append(Person(interviewee))

    interviewers = []
    for i, interviewer in enumerate(input_book.col_values(2)):
        if interviewer:
            interviewers.append(Person(interviewer, idx=i))


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
            sheet1.write(idx, 0, interviewees[i][j].name)
            idx += 1

    time_idx = 0
    while len(interviewers) != 0:
        while True:
            has_delete = False
            for i in range(len(interviewers)):
                if interviewers[i].time_idx > time_idx:
                    # 如果面试者在这个时间段已经被面过
                    continue
                if specify_final and len(interviewers[i].used_lst) == 2:
                    for j in range(len(interviewees[1])):
                        if interviewees[1][j].time_idx > time_idx:
                            # 如果面试官在这个时间段已经面了人
                            continue
                        if interviewees[1][j].name in interviewers[i].used_lst:
                            # 如果面试官已经面过这个人
                            continue
                        interviewees[1][j].time_idx = time_idx + 1
                        interviewers[i].time_idx = time_idx + 1
                        interviewers[i].used_lst[interviewees[1][j].name] = [len(interviewees[0]) + j, time_idx + 1]
                        break
                else:
                    for j in range(len(interviewees[0])):
                        if interviewees[0][j].time_idx > time_idx:
                            # 如果面试官在这个时间段已经面了人
                            continue
                        if interviewees[0][j].name in interviewers[i].used_lst:
                            # 如果面试官已经面过这个人
                            continue

                        interviewees[0][j].time_idx = time_idx + 1
                        interviewers[i].time_idx = time_idx + 1
                        interviewers[i].used_lst[interviewees[0][j].name] = [j, time_idx + 1]
                        break
                if len(interviewers[i].used_lst) == 3:
                    for k, v in interviewers[i].used_lst.items():
                        sheet1.write(v[0] + 1, v[1], interviewers[i].name)
                    interviewers.remove(interviewers[i])
                    has_delete = True
                    break
            if not has_delete:
                break
        time_idx += 1
        interviewers = sorted(interviewers)
    for i in range(1, time_idx + 1):
        sheet1.write(0, i,
                     (start_time + timedelta(minutes=(i - 1) * args.interval_time)).time().strftime('%H:%M') + '-' +
                     (start_time + timedelta(minutes=i * args.interval_time)).time().strftime('%H:%M'))
    output_book.save('面试安排.xls')
