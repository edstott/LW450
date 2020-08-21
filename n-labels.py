from LW450 import LW450

import logging
logging.basicConfig(level=logging.DEBUG,format='[%(levelname)s] (%(threadName)-10s) %(message)s')

n_labels = 20
jobs = [['{:02d}'.format(n+20)]*3 for n in range(n_labels)]
with LW450() as printer:
	for i,job in enumerate(jobs):
		printer.printText(job,dir='horizontal',labeltype='11353_right')
