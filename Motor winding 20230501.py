import numpy as np

# Initialization
S = 12                  # Slots number
P = 10                   # Pole number
Nph = 3           # Phase number

Sph = int(S/Nph)  # Slots per phase
Span = np.floor(S/P)    # Coil span
E_angle = 180*P/S       # Electrical angle per slot
Ncpp = S/P/Nph    # Slots per phase per pole

print('Nominal Coil Span =', Span)

# Calculate K0
K0 = np.zeros(int(S/2))
for q in range(int(S/2)):
    K0[q] = 2*S/3/P*(1+3*q)

K0 = np.min([int(x) for x in np.nditer(K0) if float(x).is_integer()])

print('K0 = ',K0)

# Phase A construction
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
print(A)
# Take the phase A
A = A[:,:Sph].astype(int)

print('Phase A =')
print(A)

# Phase B
B = A[2:4,:]+K0
B[B>S] = B[B>S]-S

print('Phase B =')
print(B)

# Phase C
C = B+K0
C[C>S] = C[C>S]-S

print('Phase C =')
print(C)

# Combine
ABCIn = np.vstack([A[2,:],B[0,:],C[0,:]])
ABCOut = np.vstack([A[3,:],B[1,:],C[1,:]])

# print("ABCIn = ")
# print(ABCIn)

# print("ABCOut = ")
# print(ABCOut)

# Winding
W = np.empty((S,Nph+1), dtype=object)
W[:,0] = range(1,S+1)

# Winding ABC
for j in range(1,Nph+1):
    for i in ABCIn[j-1,:]-1:
        if W[i,j] is None:
            W[i,j] = "In"
        else:
            W[i,j] += " & In"

    for i in ABCOut[j-1,:]-1:
        if W[i,j] is None:
            W[i,j] = "Out"
        else:
            W[i,j] += " & Out"

print('Winding = ')
print(W)

from openpyxl import Workbook

data = W.tolist()

# 創建 Excel
workbook = Workbook()

# 獲取當前工作表
worksheet = workbook.active

# 添加標題行
headers = ['Slot','Phase A', 'Phase B', 'Phase C']

# 將標題寫入第一行
for col, header in enumerate(headers, start=1):
    worksheet.cell(row=1, column=col, value=header)

# 將數據寫入工作表中
for row in data:
    worksheet.append(row)

# 數據保存到文件中
workbook.save('output.xlsx')