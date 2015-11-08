s='''8	Д	Н		Н(8)
8	Н		Д	
8			Н	Д
8		Д		Н
	Д	Н		
	Н		Д	
8			Н	Д
8		Д		Н
8	Д	Н		
8	Н		Д	
8			Н	Д
		Д		Н
	Д	Н		
8	Н		Д	
8			Н	Д
8		Д		Н
8	Д	Н		
8	Н		Д	
			Н	Д
		Д		Н
8	Д	Н		
8	Н		Д	
8			Н	Д
8		Д		Н
8	Д	Н		
	Н		Д	
			Н	Д
8		Д		Н
8	Д	Н		
8	Н		Д	
8			Н(4)	Д'''























ss=s.split('\n')
sss=(x.split('\t') for x in ss)

for z in zip(*sss):
    zz=''.join([x or '-' for x in z])
    print(zz)