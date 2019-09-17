def hello():
	print("Hello World!")

	filename = '../mabia_model/GrossMarginID.txt'
	infile = open(filename, 'r')
	print(infile)


a = [{'a':1}, {'a':2}, {'a':2}]

for v in a:
	if v['a']==1:
		print(v)