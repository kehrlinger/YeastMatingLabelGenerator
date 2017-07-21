# This script takes two vector and their respective inserts as arguments and
# combines them in the same way a mating assay would be carried out.
# An excel spreadsheet is created and the generated labels are saved to this spreadsheet
import openpyxl

vec1 = input('Please insert the first vector: \n(for example: pGADT7)\n')
ins_vec_1 = input ('Please enter the inserts of this vector, delimited by a whitespace: \n'
                           '(for example: insert1 insert2 insert3)\n')

vec2 = input('Please enter the second vector: \n'
             '(for example: pGBKT7)\n')
ins_vec_2 = input('Please enter the inserts of the second vector, delimited by a whitespace: \n'
                  '(for example: insert1 insert2 insert3)\n')

def construct_vectors(vector, inserts):
    vector_list = []
    for i in inserts:
        i = vector + '-' + i
        vector_list.append(i)
    return vector_list

final_ins_vec_1 = construct_vectors(vec1, ins_vec_1.split())
final_ins_vec_2 = construct_vectors(vec2, ins_vec_2.split())

labels = []
for v in final_ins_vec_1:
    for i in range(0, len(final_ins_vec_2)):
        label = v + ' x ' + final_ins_vec_2[i]
        labels.append(label)

# Construct a list with coordinates
coordinates =[]
for c in range(0, len(labels)):
    coordinate = 'A' + str(c+1)
    coordinates.append(coordinate)

wb = openpyxl.load_workbook('MatingLabels.xlsx')
sheet = wb.create_sheet(index=0, title='Labels')
for index in range(0, len(labels)):
    sheet[coordinates[index]] = labels[index]
wb.save('MatingLabels.xlsx')