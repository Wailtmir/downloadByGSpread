import datetime
import re
import numba


def dateParse(s):
	forma = r'[;,.,,,\s,\\,/,-,_+,]'
	s = re.split(forma,s)
	try:
		
		s.remove('')
		s.remove("")
	except ValueError:
		pass
	if len(s)==2:
		s+=['16']
	for i in range(0,len(s)):
		try:
			s[i] = int(s[i])
		except ValueError:
			#print("errorDateFormat")
			return [0]
	return s

def inFrame(current,frame):
	current=dateParse(current)
	#print(current)
	try:
		month = current[1]
		day=current[0]
		if(frame[0][1]==frame[1][1] and frame[0][0]<=day<=frame[1][0] and frame[1][1]==month):
			#print(1)
			return True
		elif frame[0][1]<frame[1][1] and frame[0][1]<=month<frame[1][1] and frame[0][0]<=day:
			#print(2)
			return True
		elif frame[0][1]<frame[1][1] and month==frame[1][1] and day<=frame[1][0]:
		#print(3)
			return True
		# есть проблемы с интервалами
	except IndexError:
		return False
	return False