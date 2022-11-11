from math import *
from sympy import *

x = 0.9 #x轴补偿值
y = 0.9 #y轴补偿值
e = Symbol('e') #x轴补偿真值
w = Symbol('w') #y轴补偿真值

solved_value = solve(
    [(-e)*sin(w)-x,
     e*cos(w)-y],
    [e,w]
)
#e要取正值，w要取[0,2π]
for e,w in solved_value:
    if e > 0 :
        print('x轴补偿真值为:{}'.format(e))
    if w>=0 and w<=2*pi:
        print('y轴补偿真值为:{}'.format(w))
        
#将e和w也就是XY的补偿真值变换十六进制

