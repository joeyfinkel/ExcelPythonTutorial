def get_col(ws, col):
    '''
    Get all values from a cell after the heading
    '''
    values = {}
    value = []

    for c in col:
        for row in ws[c]:
            if (row.value not in get_headers(ws)): # only get values after the header
                value += [row.value]
                values[c] = value

    return values

def get_col_letter(ws):
    columns = []

    for i in range(ws.max_column):
        columns += [get_column_letter(i + 1)]

    return columns

def get_col_value(ws):
    letter = 0
    value = ''

    for i in range(ws.max_column):
        letter = get_column_letter(i + 1)
        for row in ws[letter]:
            value = row.value

    return value

def get_rows(ws):
    students ={}
    value = []
    for row in ws.iter_rows(values_only=True):
        student_id = row[0]
        student = {
            'Name': row[1],
            'Math': row[2],
            'Science': row[3],
            'English': row[4],
            'Reading': row[5]
        }

        students[student_id] = student
        return students

# for row in get_rows(ws):
# print(json.dumps(get_rows(ws)))
# hi = get_rows(ws)