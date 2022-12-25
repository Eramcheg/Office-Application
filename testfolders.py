import os

arr = os.listdir("G:/FIles/test")
for i in range( len(arr)):
    os.rename('G:/FIles/test/'+arr[i], 'G:/FIles/test/'+str(i+1)+'0000' + '.png')
print(arr)
# alphabet= ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
# alphabet2=["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z"]
# for i in (alphabet):
#     for j in alphabet:
#         alphabet2.append(i+j)
# print(alphabet2)