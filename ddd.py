c = '''
'''


SellExcel = '''
'''
b = SellExcel.splitlines()
c = c.splitlines()
for i in b:
    if c.__contains__(i):
        c.remove(i)

with open('tt.txt', 'w', encoding='utf-8') as f:
    for ele in c:
        f.write(ele+'\n')
