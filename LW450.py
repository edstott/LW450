import win32print
import win32ui
import subprocess
from PIL import Image, ImageOps
import math
import pyx
from threading import Thread
from queue import Queue
import logging
import os

DEFAULT_PRINTER_NAME = "DYMO LabelWriter 450 Turbo"
BASE_JOB_NAME = "job"
EM_TO_LINESPACE = 1.5
DEBUG_MARKERS = False
MIN_TEXT_SIZE = -5
PRINT_ENABLE = True

def encodeimfileFn(filename):
	im = Image.open(filename)
	logging.debug(im.getpixel((0,0)))
	im = im.convert('L')
	im = ImageOps.invert(im)
	im = im.convert('1')
	
	size = im.size
	#Set width
	widthbytes = math.ceil(size[0]/8)
	
	imdata = iter(im.tobytes())

	printdata = b'\x1bD'+ widthbytes.to_bytes(1,'big')
	

	#Feed for tear-off
	#printdata += b'\x1bE'

	#Write data
	#Loop over rows
	pixcount = 0
	for y in range(size[1]):
		#Syn character
		printdata += b'\x16'
		#Loop over column bytes
		for x in range(widthbytes):
			printdata += next(imdata).to_bytes(1,'big')
	return printdata

#Enumerate special function codes for printer and provide job numbers
class jobid():
	FIND_PRINTER = -1
	STOP = -2
	FEED = -3
	def __init__(self):
		self.idval = 0
	def getid(self):
		self.idval += 1
		return self.idval

#Thread to run jobs sequentially through the printer		
class clprintdaemon(Thread):
	def __init__(self, printQueue):
		Thread.__init__(self, group=None, target=None, name=None, args=(), kwargs=None, daemon=None)
		self.printQueue = printQueue

	def run(self):
		done = False
		while done is False:
			job,data = self.printQueue.get()
			if job == jobid.STOP:
				logging.debug("Job {}: STOP".format(job))
				done = True
			if job == jobid.FIND_PRINTER:
				logging.debug("Job {}: FIND_PRINTER {}".format(job,data))
				printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_NAME,'Windows NT Local Print Providor',2)
				dymos = [p for p in printers if p['pPrinterName'] == data]
				if len(dymos) == 0:
					raise LW450.NoPrinterError
				dymo = dymos[0]
				logging.debug("Printer Found")
			
			#job id >0 indicates print job
			if job > 0:
				logging.debug("Print job {}: {} bytes of data".format(job,len(data)))
				if PRINT_ENABLE:
					hPrinter = win32print.OpenPrinter(dymo['pPrinterName'])
					try:
						hJob = win32print.StartDocPrinter(hPrinter, 1, ("test of raw data", None, "RAW"))
						try:
							win32print.StartPagePrinter(hPrinter)
							win32print.WritePrinter(hPrinter, data)
							win32print.WritePrinter(hPrinter, b"\x1bE")
							win32print.EndPagePrinter(hPrinter)
						finally:
							win32print.EndDocPrinter(hPrinter)
					finally:
						win32print.ClosePrinter(hPrinter)
				
			self.printQueue.task_done()	#Used to count jobs in queue but not necessary with explicit STOP command

#Thread to convert a pdf file into printer data			
class clpdfproc(Thread):
	def __init__(self, printQueue, jobid):
		Thread.__init__(self, group=None, target=None, name=None, args=(), kwargs=None, daemon=None)
		self.printQueue = printQueue
		self.jobid = jobid
		
	def run(self):
		jobname = BASE_JOB_NAME+str(self.jobid)

		#Use ghostscript to render pdf to black and white png
		subprocess.run(['gswin64c','-sDEVICE=pngmono','-r300','-sOutputFile='+jobname+'.png','-dNOPAUSE',jobname+'.pdf','-c','quit'],stdout=subprocess.PIPE)

		#Convert png to print data
		data = encodeimfileFn(jobname+'.png')
		
		#Put data on queue for printing
		self.printQueue.put((self.jobid,data))
		logging.debug("Job {} placed in print queue".format(self.jobid))

#Class to manage the printer
class LW450:

	#Exceptions
	class Error(Exception):
		pass
		
	class NoPrinterError(Error):
		pass
		
	class TextTooSmallError(Error):
		pass
	
	#Label dimensions
	labeltypes = {
		'11353_left':{'centre':(5.3,10.35), 'size':(10.6,20.7)},
		'11353_right':{'centre':(18,10.35), 'size':(10.6,20.7)},
		'99014':{'centre':(27,50.5), 'size':(45,101)}
		}
	
	#Named font sizes
	textsizes = {
		'tiny':-4,
		'normal':0,
		'LARGE':3
		}
	
	#Latex commands for font families
	fontfamilies = {
		'tt':r"\renewcommand{\familydefault}{\ttdefault}",
		'sf':r"\renewcommand{\familydefault}{\sfdefault}",
		'rm':r"\renewcommand{\familydefault}{\rmdefault}"
	}
	
	#Start printer thread, set up queue, set up pyx
	def __init__(self,printername = DEFAULT_PRINTER_NAME,family='sf'):
		self.job = jobid()
		self.producerthreads = []
		self.initpyx(family)
		self.printQueue = Queue()
		self.printdaemon = clprintdaemon(self.printQueue)
		self.printdaemon.start()
		self.printQueue.put((jobid.FIND_PRINTER,printername))
		
	def initpyx(self,family):
		#pyx.pyxinfo()
		pyx.unit.set(defaultunit="mm")
		pyx.text.set(pyx.text.LatexRunner,texenc='utf-8')
		pyx.text.preamble(self.fontfamilies[family])
		self.pyxready = None
			
	def __enter__(self):
		return self
	
	#Exit function waits for jobs to finish
	def __exit__(self, exc_type, exc_val, exc_tb):
		logging.debug("Waiting for producer threads")
		while True in [t.is_alive() for t in self.producerthreads]:
			pass
		logging.debug("Waiting for printer thread")
		self.printQueue.put((jobid.STOP,None))
		self.printdaemon.join()
	
	#Print a pyx canvas
	def printCanvas(self,c,labeltype='11353_left'):
	
		#Get a new job code
		id = self.job.getid()
		jobname = BASE_JOB_NAME+str(id)
		
		#Get label dimensions
		label = self.labeltypes[labeltype]
		
		#Set up page dimensions
		labelbox = bbox=pyx.bbox.bbox(0,0,label['centre'][0]+label['size'][0]/2,label['centre'][1]+label['size'][1]/2)
		#Create pdf
		pg = pyx.document.page(c,fittosize=1, margin=0, bboxenlarge=0,bbox=labelbox)
		doc = pyx.document.document([pg])
		doc.writePDFfile(jobname)
		
		#Start thread for conversion to print data
		self.producerthreads.append(clpdfproc(self.printQueue,id))
		self.producerthreads[-1].start()
		
	
	#Print a text label
	def printText(self,text,dir='vertical',labeltype='11353_left',textsize='auto',linespacestretch = 1.0,align='centre'):
		#Get a new job code
		id = self.job.getid()
		jobname = BASE_JOB_NAME+str(id)
		
		#Get label dimensions
		label = self.labeltypes[labeltype]
		
		#Set up font size
		if textsize=='auto':
			tsize = 5
		elif textsize in self.textsizes:
			tsize = self.textsizes[textsize]
		else:
			tsize = textsize
		
		#Create pyx canvas
		c = pyx.canvas.canvas()
		if DEBUG_MARKERS:
			c.stroke(pyx.path.rect(0,0,1.0,1.0))
			c.stroke(pyx.path.circle(label['centre'][0],label['centre'][1],1.0))
			c.stroke(pyx.path.rect(label['centre'][0]-label['size'][0]/2,label['centre'][1]-label['size'][1]/2,label['size'][0],label['size'][1]))
		
		#Loop until text fits the label
		autofit = False
		while not autofit:
		
			#Find linespace based on em size for font
			emc = pyx.canvas.canvas()
			emc.text(0,0,'m',[pyx.text.size(tsize)])
			linespace = pyx.unit.tomm(emc.bbox().width()) * linespacestretch * EM_TO_LINESPACE
			logging.debug("Line spacing is {}".format(linespace))
			
			#Set up origin for first line and offset for next lines
			if dir == 'vertical':
				if align == 'centre':
					textloc = (label['centre'][0] - (len(text) - 1) * linespace / 2, label['centre'][1])
				else:
					textloc = (label['centre'][0] - (len(text) - 1) * linespace / 2, label['centre'][1] - label['size'][1]/2)
				lineoff = (+linespace,0)
				linelen = label['size'][1]
			else:
				if align == 'centre':
					textloc = (label['centre'][0], label['centre'][1] + (len(text) - 1) * linespace / 2)
				else:
					textloc = (label['centre'][0] + label['size'][0]/2, label['centre'][1] + (len(text) - 1) * linespace / 2)
				lineoff = (0,-linespace)
				linelen = label['size'][0]
			
			#Set up text arguments
			#textargs = [pyx.text.parbox(linelen, baseline=pyx.text.parbox.middle),pyx.text.valign.middle,pyx.text.size(tsize)]
			textargs = [pyx.text.valign.middle,pyx.text.size(tsize)]
			if dir == 'vertical':
				textargs.append(pyx.trafo.rotate(90))
			if align == 'centre':
				textargs.append(pyx.text.halign.center)
			else:
				textargs.append(pyx.text.halign.left)
			
			#Create the text
			artbox = pyx.canvas.canvas()		
			for t in text:
				txtbox = pyx.canvas.canvas()
				txtbox.text(0,0,t,textargs)
				
				#translate each line by the line origin
				artbox.insert(txtbox,[pyx.trafo.trafo().translated(*textloc)])
				textloc = (textloc[0] + lineoff[0], textloc[1] + lineoff[1])
			
			#Check if text needs to be resized
			if textsize == 'auto':
				artbb = artbox.bbox()
				logging.debug("Text size {} x {}".format(pyx.unit.tomm(artbb.width()),pyx.unit.tomm(artbb.height())))
				if pyx.unit.tomm(artbb.width()) <= label['size'][0] and pyx.unit.tomm(artbb.height()) <= label['size'][1]:
					autofit = True
				else:
					tsize -= 1
					logging.debug("Reducing size to {}".format(tsize))
					if tsize < MIN_TEXT_SIZE:
						raise self.TextTooSmallError
			else:
				autofit = True
		
		#Insert text into canvas
		c.insert(artbox)
		#Set up page dimensions
		labelbox = bbox=pyx.bbox.bbox(0,0,label['centre'][0]+label['size'][0]/2,label['centre'][1]+label['size'][1]/2)
		#Create pdf
		pg = pyx.document.page(c,fittosize=1, margin=0, bboxenlarge=0,bbox=labelbox)
		doc = pyx.document.document([pg])
		doc.writePDFfile(jobname)
		
		#Start thread for conversion to print data
		self.producerthreads.append(clpdfproc(self.printQueue,id))
		self.producerthreads[-1].start()
		
		
#print(printdata)
	
if __name__ == "__main__":
	from datetime import datetime
	logging.basicConfig(level=logging.DEBUG,format='[%(levelname)s] (%(threadName)-10s) %(message)s')
                    
	with LW450() as printer:
		#jobs = [['Test message',datetime.today().strftime('%Y-%m-%d %H:%M'),'Line 3','Line 4']]
		jobs = [['1','','1']]
		for i,job in enumerate(jobs):
			logging.debug("Entering job {}".format(i))
			printer.printText(job,dir='horizontal',labeltype='11353_right')
	logging.debug("Finished loading jobs")
	
	