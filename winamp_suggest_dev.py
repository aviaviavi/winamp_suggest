from Tkinter import *
from tkMessageBox import *
from winamp import *
import pickle
import time
import sys
import datetime
from threading import Thread
import win32com.client

"""
Winamp Music Suggester tracks what music you listen to using Winamp, and suggests what artist you may want to listen to
at any given time.

This program is an extension of Yaron Inger's Winamp.py
"""

__author__ = "Avi Press"
__email__ = "avipress@gmail.com"
__status__ = "Production"
__maintainer__ = "Avi Press"


def getArtistName(current_track):
	"""gets the artist name from CurrentPlayingTitle"""
	artist = ''
	for i in current_track:
		if i == '-':
			break
		else:
			artist = artist + i
	return artist.strip()

def listenBackground():
	App.Terminated = False
	t = Thread(target=app.listen)  #since listen sleeps, it will halt the gui unless it runs parallel
	t.start()
	return


def stopListening():
	App.Terminated = True
	app.label.configure(text="Tracking stopped")
	app.button1.configure(text="Track my Listening", command=listenBackground)


def exit():
	App.Terminated = True
	root.destroy()
	sys.exit()

class App(object):  

	Terminated = False
		
	def __init__(self, master):
		self.frame = Frame(master)
		self.label = Label(text="Welcome to the Winamp Music Suggester!")
		self.button1 = Button(self.frame, text="Track my Listening", command=listenBackground)  #generically labeled because they will be changing in function
		self.button2 = Button(self.frame, text="Suggest an Artist", command=self.suggest)
		self.quitbutton = Button(self.frame, text="Quit", command=exit)
		self.label.pack(side=TOP, pady=50)
		self.frame.pack()
		self.button1.pack(side=LEFT)
		self.button2.pack(side=LEFT)
		self.quitbutton.pack(side=LEFT)


	def listen(self):
		#borrowed code from some guy on the internet, cant get it to work with py2exe
##		winamp_found = False
##		strComputer = "."
##		while not winamp_found:
##			if win32com.client.gencache.is_readonly:
##				win32com.client.gencache.is_readonly = False
##				win32com.client.gencache.Rebuild()
##			from win32com.client.gencache import EnsureDispatch
##			self.label.configure(text="Searching for instance of Winamp")
##			objWMIService = EnsureDispatch("WbemScripting.SWbemLocator", bForDemand=0)
##			objSWbemServices = objWMIService.ConnectServer(strComputer,"root\cimv2")
##			colItems = objSWbemServices.ExecQuery("Select * from Win32_Process")
##			process_names = [process.name for process in colItems]
##			if 'winamp.exe' in process_names:
##				winamp_found = True
		w = Winamp()
		last_played = ""
		self.button1.configure(text="Stop Tracking my Listening", command=stopListening)
		self.label.configure(text="Now tracking what bands you listen to")
		while not App.Terminated:
			i = 0
			if w.getPlaybackStatus(): # only execute if theres music playing
				currently_playing = str(w.getCurrentPlayingTitle())
				current_band = getArtistName(currently_playing)
				if last_played != currently_playing:
					last_played = currently_playing
					song_changed = True
				else:  #the person has been listening to the song, add the band to that hour in database
					if song_changed:
						#load the pickled data, for reading
						pkl_file = open('winamp_data.pkl', 'rb')
						played_bands = pickle.load(pkl_file)
						try:
							played_bands[datetime.datetime.now().hour][current_band] += 1 #increment the band's count, or add it
							song_changed = False
						except KeyError:
							played_bands[datetime.datetime.now().hour][current_band] = 1
							song_changed = False
						pkl_file.close() #so we can open it for writing. back indent this block?
						#save the dictionary
						output = open('winamp_data.pkl', 'wb')
						pickle.dump(played_bands, output)
						output.close()
				# wait. only add the next song if the person has been listening to it for >= 30 seconds
			while i < 30 and not App.Terminated:
				time.sleep(1)
				i += 1

	


	def suggest(self):
		#load the data
		suggestion_made = False
		raffle_list = self.generateRaffleList()
		if not raffle_list:
			showinfo(message="Unfortunately, there isn't any data stored for your music preferences at this time of the day. Please try again later.")
			return
		while not suggestion_made:
				if not raffle_list:
					self.label.configure(text="It seems that was the last possible artist to suggest.\n Try listening to some more music so there is more data on your music tastes.")
					return
				suggestion = random.choice(raffle_list) # the suggestion is a random selection of raffle_list
				good_suggestion = askyesno(message="At this point in the day, you may enjoy listening to is " + suggestion + ".\n Is this a good suggestion?", icon="question")
				if good_suggestion:
					# w.playArtist(suggestion) wont play the songs on desktop, query comes back empty
					suggestion_made = True
					showinfo(message="Great! I'll start tracking your listening.")
					w = Winamp()
					w.playArtist(suggestion)
					global listenBackgound
					listenBackground()

				else:
					raffle_list = self.removeArtist(raffle_list, suggestion)

	def mergeDictionaries(*args):
		merged = {}
		for i in args:
			if type(i) is not App:
				data = i.items()
				for key, value in data:
					if key in merged:
						merged[key] += value
					else:
						merged[key] = value
		return merged


	def removeArtist(self, raffle_list, artist):
		try:
			while 1:
				raffle_list.remove(artist)
		except ValueError:
			return raffle_list

	def generateRaffleList(self):
			pkl_file = open('winamp_data.pkl', 'rb')
			played_bands = pickle.load(pkl_file)
			pkl_file.close()
			now = datetime.datetime.now().hour
			possible_bands = played_bands[now] #a dictionary containing this hours data
			try:
				distance = 1
				while len(possible_bands) < 1: #go +- 1 hours from now, until there are at least 1 bands to chose from
					if distance >= 12:
						return False
					possible_bands = self.mergeDictionaries(possible_bands, played_bands[(now - distance) % 24], played_bands[(now + distance) % 24])
					distance += 1
			except IndexError, KeyError:
				return False
			raffle_list = list()
			for key in possible_bands: #add the band to the list as many times as its been played
				for i in range(int(possible_bands[key])):
					raffle_list.append(key)
			return raffle_list


if __name__ == '__main__':
	try:
		pkl_file = open('winamp_data.pkl', 'rb')
		pkl_file.close()
	except IOError: #first time program has been used, initialize listening history
		played_bands = [] # list of dictionaries, one dictionary per hour
		for i in range(24):
			played_bands.append({})
			#save the dictionary via pickle
			output = open('winamp_data.pkl', 'wb')
			pickle.dump(played_bands, output)
			output.close()
	finally:
		root = Tk()
		app = App(root)
		root.title("Music Suggest")
		root.mainloop()
