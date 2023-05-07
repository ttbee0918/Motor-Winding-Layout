import numpy as np
S = 15
P = 4
Phase_Num = 3   # phase number\

Nph = S/Phase_Num
Span = np.floor(S/P)
E_angle = 180*P/S
Ncpp = S/P/Phase_Num

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
B = int(S/Phase_Num)
A = A[:,:B]

print(A)

# Calculate K0
K0 = np.zeros(int(S/2))
for q in range(int(S/2)):
    K0[q] = 2*S/3/P*(1+3*q)

print(K0)

K0 = [int(x) for x in np.nditer(K0) if float(x).is_integer()]

print(K0)