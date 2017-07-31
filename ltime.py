from __future__ import print_function
import time
import sys

class LTime(object):
	"""docstring for Capture"""
	def __init__(self):
		super(Capture, self).__init__()
	
	@staticmethod
	def delay(wtime=5):
		print('Sleeping %d: '%wtime, end="")
		for i in range(wtime):
			print('.', end="")
			time.sleep(1)
		print("")

if __name__ == '__main__':
	LTime.delay(10)

	