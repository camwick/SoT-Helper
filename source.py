import pandas as pd 
from tkinter import *
from PIL import ImageTk, Image
import webbrowser

def main():
	# setting up gui
	root = Tk()
	root.title('SoT Helper')
	root.iconbitmap('./images/logo.ico')
	root.geometry('400x100')

	# reading excel file
	df_outposts = pd.read_excel('SoT_Info.xlsx', sheet_name='Outposts')
	islands = df_outposts['Island']
	df_tales = pd.read_excel('SoT_Info.xlsx', sheet_name='Tales')
	tale_names = df_tales['Tale']

	# creating menu bar
	menubar = Menu(root)
	outpost_menu = Menu(menubar, tearoff=0)
	menubar.add_cascade(label='Outposts', menu=outpost_menu)
	tales_menu = Menu(menubar, tearoff=0)
	menubar.add_cascade(label='Tall Tales', menu=tales_menu)
	shores_menu = Menu(menubar, tearoff=0)
	tales_menu.add_cascade(label='Shores of Gold Arc', menu=shores_menu)

	# creating main frame
	frame = Frame(root, padx=10, pady=10)
	frame.pack()

	# direction label
	info1 = Label(frame, text='Make selection by using the', font=(None, 12))
	info2 = Label(frame, text=' above menus', font=(None, 12))
	info1.grid(row=0, column=0, columnspan=3)
	info2.grid(row=1, column=0, columnspan=3)

	# creating options in outpost menu
	# command first deletes everything in widget then builds frame
	outpost_menu.add_command(label=islands[0], command=lambda: [delete(frame), outpost(root, frame, df_outposts.loc[0])])
	outpost_menu.add_command(label=islands[1], command=lambda: [delete(frame), outpost(root, frame, df_outposts.loc[1])])
	outpost_menu.add_command(label=islands[2], command=lambda: [delete(frame), outpost(root, frame, df_outposts.loc[2])])
	outpost_menu.add_command(label=islands[3], command=lambda: [delete(frame), outpost(root, frame, df_outposts.loc[3])])
	outpost_menu.add_command(label=islands[4], command=lambda: [delete(frame), outpost(root, frame, df_outposts.loc[4])])
	outpost_menu.add_command(label=islands[5], command=lambda: [delete(frame), outpost(root, frame, df_outposts.loc[5])])
	outpost_menu.add_command(label=islands[6], command=lambda: [delete(frame), outpost(root, frame, df_outposts.loc[6])])

	# creating options in Tall Tales menu
	# command first deletes everything in widget then builds frame
	shores_menu.add_command(label=tale_names[0], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[0])])
	shores_menu.add_command(label=tale_names[1], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[1])])
	shores_menu.add_command(label=tale_names[2], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[2])])
	shores_menu.add_command(label=tale_names[3], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[3])])
	shores_menu.add_command(label=tale_names[4], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[4])])
	shores_menu.add_command(label=tale_names[5], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[5])])
	shores_menu.add_command(label=tale_names[6], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[6])])
	shores_menu.add_command(label=tale_names[7], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[7])])
	shores_menu.add_command(label=tale_names[8], command=lambda: [delete(frame), talltales(root, frame, df_tales.loc[8])])

	# builds gui
	root.config(menu=menubar)
	root.mainloop()

# deletes everything in the frame
def delete(frame):
	for widget in frame.winfo_children():
		widget.destroy()

# builds the outposts framespip
def outpost(root, frame, row):
	# changing window size
	root.geometry('400x865')

	# get map information
	df_map = pd.read_excel('SoT_Info.xlsx', sheet_name='Map Location')
	mapLocation = df_map[row['Island']]

	# creating title
	title = Label(frame, font=(None,18), text=row['Island'] + f' ({mapLocation[0]})')
	
	# load image
	img = ImageTk.PhotoImage(Image.open('./images/Outposts/' + row[1] + '.png'))
	img_lbl = Label(frame, image=img)
	img_lbl.photo = img

	# region
	region = Label(frame, text='Region: ' + row['Region'], font=(None, 16))

	# Emissaries
	emissaries = Label(frame, text='---- Emisarries ----', font=(None, 16))
	gold = Label(frame, text='- Gold Hoarder: ' + row['Gold'], font=(None, 14))
	merchant = Label(frame, text='- Merchant Alliance: ' + row['Merchant'], font=(None, 14))
	souls = Label(frame, text='- Order of Souls: ' + row['Souls'], font=(None, 14))

	# shop keepers
	shops = Label(frame, text='---- Shop Keepers ----', font=(None, 16))
	tavern = Label(frame, text='- Tavern: ' + row['Tavern'], font=(None, 14))
	weapon = Label(frame, text="- Weaponsmith's Shop: " + row['Weapon'], font=(None, 14))
	equipment = Label(frame, text='- Equipment Shop: ' + row['Equipment'], font=(None, 14))
	clothing = Label(frame, text='- Clothing Shop: ' + row['Clothing'], font=(None, 14))
	shipwright = Label(frame, text='- Shipwright Shop: ' + row['Shipwright'], font=(None, 14))
	emporium = Label(frame, text='- Pirate Emporium: ' + row['Emporium'], font=(None, 14))

	# empty rows for spacings
	space1 = Label(frame, text=' ')
	space2 = Label(frame, text=' ')
	space3 = Label(frame, text=' ')

	# adds special name for the last Tall Tale
	if row['Special'] == 'Grace':
		space3.grid(row=16, column=0, columnspan=3)
		special = Label(frame, text='---- Special ----', font=(None, 16))
		alliance = Label(frame, text='- Alliance Leader: ' + row['Special'], font=(None, 14))
		special.grid(row=17, column=0, columnspan=3)
		alliance.grid(row=18, column=0, columnspan=3)

	# putting everthing on screen
	title.grid(row=0, column=0, columnspan=3)
	img_lbl.grid(row=1, column=0, columnspan=3)
	region.grid(row=2, column=0, columnspan=3)
	space1.grid(row=3, column=0, columnspan=3)
	emissaries.grid(row=4, column=0, columnspan=3)
	gold.grid(row=5, column=0, columnspan=3)
	merchant.grid(row=6, column=0, columnspan=3)
	souls.grid(row=7, column=0, columnspan=3)
	space2.grid(row=8, column=0, columnspan=3)
	shops.grid(row=9, column=0, columnspan=3)
	tavern.grid(row=10, column=0, columnspan=3)
	weapon.grid(row=11, column=0, columnspan=3)
	equipment.grid(row=12, column=0, columnspan=3)
	clothing.grid(row=13, column=0, columnspan=3)
	shipwright.grid(row=14, column=0, columnspan=3)
	emporium.grid(row=15, column=0, columnspan=3)

# builds shores of gold arc frames
def talltales(root, frame, row):
	root.geometry('850x550')

	# title
	title = Label(frame, text=row['Tale'], font=(None, 18), padx=10, pady=10)

	# section labels
	location_info = Label(frame, text='Location:', font=(None, 18), padx=10, pady=10)
	journal_info = Label(frame, text='Journal Information:', font=(None, 18), padx=10, pady=10)

	# loading starting location image
	img = ImageTk.PhotoImage(Image.open('./images/Start Locations/' + row['Start'] + '.png'))
	img_lbl = Label(frame, image=img)
	img_lbl.photo = img

	# start location tag
	if row['Start'] == 'tavern':
		tale_start = Label(frame, text='Tall Tale Starts at: any ' + row['Start'], font=(None,16))
	else:	
		tale_start = Label(frame, text='Tall Tale Starts at: ' + row['Start'], font=(None,16))
	tale_person = Label(frame, text='Near ' + row['Person'], font=(None,16))	

	# adding hyperlinks for more info
	link1 = Label(frame, text='Interactive Map', fg='blue', cursor='hand2', font=(None,16))
	link2 = Label(frame, text='Wiki', fg='blue', cursor='hand2', font=(None,16))

	# adding picture viewer
	if row['Tale'] != 'Shores of Gold':
		# loading pics and passing them to image_viewer
		img1 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location1'] + '.png'))
		img2 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location2'] + '.png'))
		img3 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location3'] + '.png'))
		img4 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location4'] + '.png'))
		img5 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location5'] + '.png'))
		img1.photo = img1
		img2.photo = img2
		img3.photo = img3
		img4.photo = img4
		img5.photo = img5

		image_list = [img1, img2, img3, img4, img5]

		image_viewer(frame, row, image_list, 5)
	else:
		# loading pics and passing them to image_viewer
		img1 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location1'] + '1.png'))
		img2 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location2'] + '2.png'))
		img3 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location3'] + '3.png'))
		img4 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location4'] + '4.png'))
		img5 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location5'] + '5.png'))
		img6 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location6'] + '6.png'))
		img7 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location7'] + '7.png'))
		img8 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location8'] + '8.png'))
		img9 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location9'] + '9.png'))
		img10 = ImageTk.PhotoImage(Image.open('./images/Journal Locations/' + row['Location10'] + '10.png'))
		img1.photo = img1
		img2.photo = img2
		img3.photo = img3
		img4.photo = img4
		img5.photo = img5
		img6.photo = img6
		img7.photo = img7
		img8.photo = img8
		img9.photo = img9
		img10.photo = img10

		image_list = [img1, img2, img3, img4, img5, img6, img7, img8, img9, img10]

		image_viewer(frame, row, image_list, 10)	

	# putting stuff on screen
	title.grid(row=0, column=0, columnspan=4)
	location_info.grid(row=1, column=0, columnspan=2)
	journal_info.grid(row=1, column=2, columnspan=3)
	img_lbl.grid(row=2, column=0, columnspan=2)
	tale_start.grid(row=3, column=0, columnspan=2)
	tale_person.grid(row=4, column=0, columnspan=2)
	link1.grid(row=5,column=0)
	link2.grid(row=5, column=1)

	# calling callback function to add links to the hyperlinks
	link1.bind('<Button-1>', lambda e: callback("https://maps.seaofthieves.rarethief.com"))
	link2.bind('<Button-1>', lambda e: callback("https://seaofthieves.gamepedia.com/Tall_Tales#Shores_of_Gold_Arc"))
	
def image_viewer(frame, row, image_list, amount):
	# loading in map info 
	df_map = pd.read_excel('SoT_Info.xlsx', sheet_name='Map Location')
	map = df_map[row['Location1']]

	# load first image
	img = image_list[amount-amount]
	img = Label(frame, image=img)
	img.grid(row=2, column=2, columnspan=3)

	# journal number counter
	img_counter = Label(frame, text=f'Journal 1 of {amount}', font=(None,14))
	img_counter.grid(row=3, column=3)

	# image island name
	img_lbl = Label(frame, text='Island: ' + row['Location1'] + f' ({map[0]})', font=(None,14))
	img_lbl.grid(row=4, column=2, columnspan=3)

	# journal name
	journal_name = Label(frame, text='Name of Journal: ' + row['Journal1'], font=(None,14))
	journal_name.grid(row=5, column=2, columnspan=3)

	# buttons
	back_btn = Button(frame, text='<<', state=DISABLED)
	forward_btn = Button(frame, text='>>', command=lambda: forward(frame, journal_name, img, img_counter, img_lbl, row, df_map, image_list, 2, amount))

	back_btn.grid(row=3, column=2)
	forward_btn.grid(row=3, column=4)

def back(frame, journal_name, img, img_counter, img_lbl, row, df_map, image_list, new, max):
	# forgetting grid items to update them
	img.grid_forget()
	img_counter.grid_forget()
	img_lbl.grid_forget()
	journal_name.grid_forget()

	# grabbing map info for correct row
	map = df_map[row['Location' + f'{new}']]

	# rewriting the grids with correct information
	img = Label(frame, image=image_list[new-1])
	img.grid(row=2, column=2, columnspan=3)

	img_counter = Label(frame, text=f'Journal {new} of {max}', font=(None,14))
	img_counter.grid(row=3, column=3)

	img_lbl = Label(frame, text='Island: ' + row[f'Location{new}'] + f' ({map[0]})', font=(None,14))
	img_lbl.grid(row=4, column=2, columnspan=3)

	journal_name = Label(frame, text='Name of Journal: ' + row[f'Journal{new}'], font=(None,14))
	journal_name.grid(row=5, column=2, columnspan=3)

	# new buttons overlapping old ones
	forward_btn = Button(frame, text='>>', command=lambda: forward(frame, journal_name, img, img_counter, img_lbl, row, df_map, image_list, new+1, max))
	back_btn = Button(frame, text='<<', command=lambda: back(frame, journal_name, img, img_counter, img_lbl, row, df_map, image_list, new-1, max))

	if new == 1:
		back_btn = Button(frame, text='<<', state=DISABLED)

	forward_btn.grid(row=3, column=4)
	back_btn.grid(row=3, column=2)

def forward(frame, journal_name, img, img_counter, img_lbl, row, df_map, image_list, new, max):
	# forgetting grid items to update them
	img.grid_forget()
	img_counter.grid_forget()
	img_lbl.grid_forget()
	journal_name.grid_forget()

	# grabbing map info for correct row
	map = df_map[row['Location' + f'{new}']]

	# rewriting the grids with correct information
	img = Label(frame, image=image_list[new-1])
	img.grid(row=2, column=2, columnspan=3)

	img_counter = Label(frame, text=f'Journal {new} of {max}', font=(None,14))
	img_counter.grid(row=3, column=3)

	img_lbl = Label(frame, text='Island: ' + row[f'Location{new}'] + f' ({map[0]})', font=(None,14))
	img_lbl.grid(row=4, column=2, columnspan=3)

	journal_name = Label(frame, text='Name of Journal: ' + row[f'Journal{new}'], font=(None,14))
	journal_name.grid(row=5, column=2, columnspan=3)

	# new buttons overlapping old ones
	forward_btn = Button(frame, text='>>', command=lambda: forward(frame, journal_name, img, img_counter, img_lbl, row, df_map, image_list, new+1, max))
	back_btn = Button(frame, text='<<', command=lambda: back(frame, journal_name, img, img_counter, img_lbl, row, df_map, image_list, new-1, max))

	if new == max:
		forward_btn = Button(frame, text='>>', state=DISABLED)

	forward_btn.grid(row=3, column=4)
	back_btn.grid(row=3, column=2)

# opens link when clicked
def callback(url):
	webbrowser.open_new(url)

# main function initiator
if __name__ == '__main__':
	main()