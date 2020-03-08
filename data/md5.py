import hashlib

def md5(ss):  
	m = hashlib.md5()
	m.update(ss)
	return m.hexdigest()
	
print md5('meimima123')