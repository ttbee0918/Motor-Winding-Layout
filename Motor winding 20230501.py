import numpy as np
S = 15                  # Slots number
P = 4                   # Pole number
Phase_Num = 3           # Phase number

Nph = int(S/Phase_Num)  # Slots per phase
Span = np.floor(S/P)    # Coil span
E_angle = 180*P/S       # Electrical angle per slot
Ncpp = S/P/Phase_Num    # Slots per phase per pole

# Initialization
A = np.zeros((4,S))
A[0,:] = np.arange(1,S+1).reshape(-1,1).ravel()
A[1,:] = E_angle*np.arange(0,S).reshape(-1,1).ravel()
A[2,:] = np.arange(1,S+1).reshape(-1,1).ravel()
A[3,:] = np.arange(1+Span,S+1+Span).reshape(-1,1).ravel()

for i in range(len(A[3,:])):
    if A[3,i] > S:
        A[3,i] -= S

# rerange 180 ~ -180
A[1,:] = np.mod(A[1,:]+180,360)-180

# reverse >90 or <-90
for i in range(len(A[0,:])):
    if A[1,i] > 90:
       A[1,i] = A[1,i] -180
       A[2,i] = A[3,i]
       A[3,i] = A[0,i]

    if A[1,i] < -90:
       A[1,i] = A[1,i] +180
       A[2,i] = A[3,i]
       A[3,i] = A[0,i]

# Sort the 1 column
A = A[:,np.argsort(np.abs(A[1]),kind='mergesort')]

# Take the phase A
A = A[:,:Nph].astype(int)

print(A)

# Winding A
W = np.empty((S,3), dtype=object)
W[A[2,:]-1,0] = 'In'
W[A[3,:]-1,0] = 'Out'
print(W)


# Calculate K0
K0 = np.zeros(int(S/2))
for q in range(int(S/2)):
    K0[q] = 2*S/3/P*(1+3*q)

K0 = np.min([int(x) for x in np.nditer(K0) if float(x).is_integer()])

print(K0)

from openpyxl import Workbook

data = A.tolist()

# 創建 Excel
workbook = Workbook()

# 獲取當前工作表
worksheet = workbook.active

# 添加標題行
headers = ['Coil','Angle', 'In', 'Out']

# 將標題寫入第一行
for col, header in enumerate(headers, start=1):
    worksheet.cell(row=1, column=col, value=header)

# 將數據寫入工作表中
for row in data:
    worksheet.append(row)

# 數據保存到文件中
workbook.save('output.xlsx')