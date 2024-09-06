#criamos nosso dicionario
d={}
a = "sim"
#utilizei o lower pra sempre deixar a resposta em minusculo, pra nao ter problemas com input maiusculo
while a.lower() == "sim":
    item = input("Digite o item que deseja adicionar: ")
    valor = float(input ("Digite o valor do item %s: "%item))
    d[item] = valor
    a = input ("Deseja adicionar mais algum item? sim/não: ")
#para exibir em for, utilizamos a seguinte sintaxe
for x in d:
    #como o x vai rodar sempre com o valor da chave (nome), podemos chamar o valor do preço com a o dicionario[x] (d[x])
    print ("Item: %s \tPreço: %.2f"%(x, d[x]))