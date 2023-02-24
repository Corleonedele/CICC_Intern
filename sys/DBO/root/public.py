



class Public():

    def writeRow(st, row_index, write_info, start=0, end=0):   
        if end == 0:
            end = st.max_column
        for index, i in enumerate(write_info):
            st.cell(row_index, index+start).value = i


    def writeColumn(st, column_index, write_info, start=0, end=0):   
        if end == 0:
            end = st.max_row
        for index, i in enumerate(write_info):
            st.cell(index+start, column_index).value = i


    def readRow(st, row_index, start=0, end=0):
        result = []
        if end == 0:
            end = st.max_column
        for i in range(start, end+1):
            result.append(st.cell(row_index, i).value)
        return result

    def readColumn(st, column_index, start=0, end=0):
        result = []
        if end == 0:
            end = st.max_row
        for i in range(start, end+1):
            result.append(st.cell(i, column_index).value)
        return result