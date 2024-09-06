#Megasena

import random

x = int(input('Digite o número de jogos a serem feitos: '))

for i in range(x):
    ran = range(1,60)
    y = random.sample(ran,6)
    y.sort()
    print(y)
input('')            