import jinja2
from jinja2 import Template


def main(s):
    tpl = Template(s)
    print(tpl.render())

table = [['我们来插入一个表格', '我们来插入一个表格', '我们来插入一个表格', '我们来插入一个表格'],
['这是一级标题1', '这是二级标题1.1', '这是三级标题1.1.1', '总结'],
['这是一级标题1', '这是二级标题1.1', '这是三级标题1.1.2', '总结'],
['这是一级标题1', '这是二级标题1.1', '这是三级标题1.1.3', '总结'],
['这是一级标题1', '这是二级标题1.2', '这是三级标题1.2.1', '总结'],
['这是一级标题1', '这是二级标题1.3', '这是三级标题1.3.1', '总结'],
['这是一级标题1', '这是二级标题1.3', '这是三级标题1.3.2', '总结'],
['别忙，还有内容', '别忙，还有内容', '别忙，还有内容', '别忙，还有内容'],
['内容', '内容', '另一段内容', '另一段内容']]

def _table_matrix():
    if not table:
        return ""

    # 处理同一行的各列
    temp_matrix = []
    for row in table:
        if not row:
            continue

        col_last = [row[0], 1, 1]
        line = [col_last]
        for i, j in enumerate(row):
            if i == 0:
                continue

            if j == col_last[0]:
                col_last[2] += 1
                line.append(["", 0, 0])
            else:
                col_last = [j, 1, 1]
                line.append(col_last)

        temp_matrix.append(line)

    # 处理不同行
    matrix = [temp_matrix[0]]
    last_row = []
    for i, row in enumerate(temp_matrix):
        if i == 0:
            last_row.extend(row)
            continue

        new_row = []
        for p, r in enumerate(row):
            if p >= len(last_row):
                break

            last_pos = last_row[p]

            if r[0] == last_pos[0] and last_pos[0] != "":
                last_row[p][1] += 1
                new_row.append(["", 0, 0])
            else:
                last_row[p] = row[p]
                new_row.append(r)

        matrix.append(new_row)

    return matrix


def table2html(t):
    # table = _fill_blank(t)
    matrix = _table_matrix()
    print(matrix)
    print("------------------------")

    html = ""
    for row in matrix:
        tr = "<tr>"
        for col in row:
            if col[1] == 0 and col[2] == 0:
                continue

            td = ["<td"]
            if col[1] > 1:
                td.append(" rowspan=\"%s\"" % col[1])
            if col[2] > 1:
                td.append(" colspan=\"%s\"" % col[2])
            td.append(">%s</td>" % col[0])

            tr += "".join(td)
        tr += "</tr>"
        html += tr

    return html


def _fill_blank(table):
    cols = max([len(i) for i in table])

    new_table = []
    for i, row in enumerate(table):
        new_row = []
        [new_row.extend([i] * int(cols / len(row))) for i in row]
        print(new_row)
        new_table.append(new_row)
    print("-----------------")
    return new_table


print(table2html(table))
# main({{ table|safe }})