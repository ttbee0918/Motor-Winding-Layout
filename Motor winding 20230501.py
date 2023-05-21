import numpy as np

# Initialization
S = 12      # Slots number
P = 10      # Pole number
Nph = 3     # Phase number

Ncph = int(S/Nph)       # Slots per phase
Span = np.floor(S/P)    # Coil span
E_angle = 180*P/S       # Electrical angle per slot
Ncpp = S/P/Nph          # Slots per phase per pole

print('Nominal Coil Span =', Span)

# Calculate K0
K0 = np.zeros(int(S/2))
for q in range(int(S/2)):
    K0[q] = 2*S/Nph/P*(1+Nph*q)
print('Possible K0 =')
print(K0)

# Choose the smallest as K0
K0 = np.min([int(x) for x in np.nditer(K0) if float(x).is_integer()])
print('Choose K0 = ',K0)

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

# Sort the phase offset
A = A[:,np.argsort(A[1],kind='mergesort')]

# Sort the abs phase offset
A = A[:,np.argsort(np.abs(A[1]),kind='mergesort')]

print('All possible coils for phase A =')
print(A)

# Take the phase A
A = A[:,:Ncph].astype(int)

print('Phase A =')
print(A)

# Phase Matrix ABC
ABC = np.tile(A[2:4,:], (Nph, 1))

# Phase Matrix Arrange
for i in range(0,Nph-1):
    ABC[2*i+2:2*i+4,:] = ABC[2*i+0:2*i+2,:]+K0
    ABC[ABC>S] = ABC[ABC>S]-S

# Winding
W = np.empty((S,Nph+1), dtype=object)
W[:,0] = range(1,S+1)

# Winding ABC
for j in range(1,Nph+1):
    for i in ABC[2*j-2,:]-1:
        if W[i,j] is None:
            W[i,j] = "In"
        else:
            W[i,j] += " & In"

    for i in ABC[2*j-1,:]-1:
        if W[i,j] is None:
            W[i,j] = "Out"
        else:
            W[i,j] += " & Out"

print('Winding Layout = ')
print(W)

from openpyxl import Workbook

# 創建 Excel
workbook = Workbook()

# 獲取當前工作表
worksheet = workbook.active

# 將數據寫入 Excel
worksheet.append(['Slot',S])
worksheet.append(['Pole',P])
worksheet.append(['Phase number',Nph])
worksheet.append(['Slots per phase',Ncph])
worksheet.append(['Coil span',Span])
worksheet.append(['Electrical angle per slot',E_angle])
worksheet.append(['Slots per phase per pole',Ncpp])

# Winding 寫入 Excel
data = W.tolist()
headers = ['Slot','Phase A', 'Phase B', 'Phase C']
worksheet.append(headers)
for row in data:
    worksheet.append(row)

# 數據保存到文件中
workbook.save('output.xlsx')