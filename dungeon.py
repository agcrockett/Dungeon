from xlrd import open_workbook
import time
import copy

#import msvcrt
#import xlsxwriter

'''
Issues:
inventory not functional
combat not interesting
'''

############################################################################################
# Pull Functions 
############################################################################################	

# this function pulls all the objects in the file "dungeon_objects.xlsx"
def pull_all_objects():
	global global_dict
	
	# create a list of the dictionaries (and workbook pages) that will be used
	dict_list = ['interactables_dict', 'consumables_dict', 'monsters_dict', 'player_info_dict']
	
	# initialise a common dictionary for these dictionaries
	global_dict = {}
	
	# initialise empty dictionaries for each type specified above
	for dict in dict_list:
		global_dict[dict] = ''
	
	# open the excel file
	book = open_workbook('dungeon_objects.xlsx')
	
	# iterate through the sheets in the excel file
	for sheet_number in xrange(0, book.nsheets):
	
		# go to the first sheet (0)
		current_sheet = book.sheet_by_index(sheet_number)
		
		# get object names (and set as dictionary keys) for each object type from row 1 
		# (first named object) to the bottom of the specified sheet
		keys = [current_sheet.cell(row_index, 0).value for row_index in xrange(1, \
		current_sheet.nrows)]
		
		# initialise the dictionary entry for this sheet, e.g. [dict_list][0] = 'interactables_list'
		global_dict[dict_list[sheet_number]] = {}
		
		# initialise the dictionary entry for each object retrieved from the sheet
		for key in keys:
			global_dict[dict_list[sheet_number]][key] = {}

		keys_index = 0

		# iterate through each row and column, adding information to the global dictionary for each
		# object
		for row_index in xrange(1, current_sheet.nrows):		
			for col_index in xrange(0, current_sheet.ncols):
				new_value = {}
				new_value = {current_sheet.cell(0, col_index).value: current_sheet.cell\
				(row_index, col_index).value}
				global_dict[dict_list[sheet_number]][keys[keys_index]].update(new_value)
			keys_index += 1
			

'''
# Allow user to input a name and gender	
def personalisation():
	global player_name
	
	player_name = raw_input("\nWhat is your name?\n> ")
		
	while True:
		player_gender = None
		gender_options = ['male', 'female']
		
		if player_gender in gender_options:
			print player_gender
		else:
			player_gender = raw_input("\nAre you male or female?\n> ").lower()
		
		break
	
	global_dict['player_info_dict'][player_name] = global_dict['player_info_dict'].pop('')
	global_dict['player_info_dict'][player_name]['name'] = player_name
	global_dict['player_info_dict'][player_name]['gender'] = player_gender
'''
############################################################################################
# Classes 
############################################################################################	
		
# Instantiate a room with described attributes
class Dungeon_Room:
	
	roomCount = 0
		
	def __init__(self, description, consumables, monsters, interactables, room_number, modified):
		self.description = description
		self.consumables = consumables
		self.monsters = monsters
		self.interactables = interactables
		self.room_number = room_number
		self.modified = modified
		
		Dungeon_Room.roomCount += 1

# Instantiate the player character with described attributes
class Player:
	def __init__(self, name, gender, description, inventory, player_hp, player_mp, \
	mood, level, equipped, room_number, str, dex, con, int, alive):
		self.name = name
		self.gender = gender
		self.description = description
		try:
			self.inventory = {inventory : current_room_objects[inventory]}
		except:
			self.inventory = {}
		self.player_hp = player_hp
		self.player_mp = player_mp
		self.level = level
		self.equipped = equipped
		self.mood = mood
		self.room_number = room_number
		self.str = str
		self.dex = dex
		self.con = con
		self.int = int
		self.alive = alive
	
	def self_description(self):
		# inv_items = self.inventory.keys()
				
		return "Your name is %s. You are holding a %s, and you feel %s."\
		% (self.name, self.equipped, self.mood)
		
# Instantiate a consumable with described attrs
class Consumable:
	
	def __init__(self, name, type, effect, description, consumed, \
	consumed_description, room_number, takeable, discovered, possessed, location):
		self.name = name
		self.type = type
		self.effect = effect
		self.description = description
		self.consumed = consumed
		self.consumed_description = consumed_description
		self.room_number = room_number
		self.takeable = takeable
		self.discovered = discovered
		self.possessed = possessed
		self.location = location

# Instantiate a monster with described attrs	
class Monster:
	
	def __init__(self, name, description, dead_description, type, hp, mp, level, items, mood, \
	room_number, str, dex, con, int, alive):
		self.name = name
		self.description = description
		self.dead_description = dead_description
		self.type = type
		self.hp = hp
		self.mp = mp
		self.level = level
		self.items = items
		self.mood = mood
		self.room_number = room_number
		self.str = str
		self.dex = dex
		self.con = con
		self.int = int
		self.alive = alive

# Instantiate an interactable of the dungeon with following attributes:
#(self, name, type, interact_effect, description, usable, openable, opened, room_number)	
class Interactable:
	
	def __init__(self, name, type, interact_effect, description, usable, openable, opened, locked,\
	room_number, contents, takeable, discovered, possessed, location):
		self.name = name
		self.type = type
		self.interact_effect = interact_effect
		self.description = description
		self.usable = usable
		self.openable = openable
		self.opened = opened
		self.locked = locked
		self.room_number = room_number
		self.contents = {}
		self.takeable = takeable
		self.discovered = discovered
		self.possessed = possessed
		self.location = location
		
############################################################################################
# Actions 
############################################################################################	
	
# opens an openable object - doors and roum_counter not working yet
def open(object):
	current_object = current_room_objects[object]
		
	if hasattr(current_object, 'openable') and current_object.openable == True:
		if current_object.opened == False:
			if current_object.locked == True:
				print "It's locked." 
				
			else: 
				print current_object.interact_effect
				current_object.opened = True
				
				if current_object.type == 'chest':
					# need to instantiate whatever object was contained in the interactable_dict contents for the chest
					contents = current_room_objects[global_dict['interactables_dict']['chest']['contents']]
					# add the contents of the chest to its initial instance's contents
					current_object.contents[contents.name] = contents
					
					if is_empty(current_object.contents) == False:
						# need to add for loop for multiple objects
						print "\nInside you can make out %s." % contents.description
						contents.takeable = True
						
					elif is_empty(current_object.contents) == True:
						print "There's nothing inside. Locator 1: open."
						
					else:
						print "Something went wrong."
						
				elif current_object.type == 'door':
					pass
					
				else:
					print "Something went wrong. Locator 2: open."
					
		else:
			print "It's already open. Locator 3: open"
	
	else:
		print "You can't %s that. Locator 4: Open." % player_action

# consume an object		
def consume(action, object):
	
	current_object = current_room_objects[object]
	
	if isinstance(current_object, Consumable):
		if current_object.consumed == False:
			if current_object.type == 'food' and action == 'eat':
				player_text = "health"
				if player_character.player_hp <= (100 - int(current_object\
				.effect)):
					print "You %s the %s. %s %d." % (action, object, player_text, \
					(current_object.effect + player_character.player_hp))
					player_character.player_hp += 5
					current_object.consumed = True
				else:
					print "You don't need to %s that. You are at full %s." % (action, player_text)
		
			elif current_object.type == 'drink' and action == 'drink':
				player_text = "mana"
				if player_character.player_mp <= (100 - int(current_object.effect)):
					print "You %s the %s. %s %d." % (action, object, player_text, \
					(current_object.effect + player_character.player_mp))
					player_character.player_mp += 5
					current_object.consumed = True
				else:
					print "You don't need to %s that. You are at full %s." % (action, player_text)
			else:
				print "That's impossible. Locator 1: consume."
		
		else:
			print "It's all gone. Locator 2: consume."
		
	else:
		print "You can't do that. Locator 3: consume."
		return 

# use an object - not working	
def use(object):
	global next_room
	
	inv_or_room(object)
	
	if hasattr(current_object, 'usable') and current_object.usable == True:
		if current_object.usable == True and not current_object.type == 'door':

			if current_object.type != 'key':
				print current_object.interact_effect
				print "Locator 1: use."
				
			elif current_object.type == 'key':
				print "Use key on what?"
				use_loop(current_object)
				
		elif current_object.locked == False and current_object.type == 'door':
			doors = dungeon_map[current_room.room_number].keys()
			
			if len(doors) > 1:
				print "Which door would you like to use?"
				use_loop(current_object)
				
			elif len(doors) <= 1:
				if dungeon_map[current_room.room_number].has_key('north'):
					next_room = dungeon_map[current_room.room_number]['north']
					print "You pass through the doorway to the north."
					
				if dungeon_map[current_room.room_number].has_key('south'):
					next_room = dungeon_map[current_room.room_number]['south']
					print "You pass through the doorway to the south."
					
			room_change()
				
		else:
			print "You can't %s that. Locator 2: use." % player_action
	elif object == 'door' and current_object.locked == True:
		print "It's locked."
	else:
		print "You can't %s that. Locator 3: use." % player_action

# examine a given object
def examine(object):

	inv_or_room(object)

	if object in current_room_objects or object in player_character.inventory:
	
		# consumables have different descriptions if consumed or not
		if isinstance(current_object, Consumable) and \
		current_object.consumed == False:
			print "You see %s." % current_object.description
		
		elif isinstance(current_object, Consumable) and \
		current_object.consumed == True:
			print "All that's left is %s." % current_object.consumed_description
		
		# room has a different description to other objects
		elif current_object.type == 'room':
			print current_room.description
		
		# if object is an interactable check for closed or opened
		elif isinstance(current_object, Interactable):
		
			if current_object.openable == False:
				print "You see %s." % current_object.description
				
			elif current_object.openable == True:
			
				if current_object.opened == True:
					if current_object.type == 'chest':
						print "You see %s. It is open." % (current_object.description)
						if current_object.contents.viewvalues() != None:
							# need to add for loop for multiple objects
							contents = current_room_objects[current_object.contents].description
							print "\nInside you can make out %s." % contents	
						else:
							print "There is nothing inside."
						
					elif current_object.type == 'door':
						print "You see %s. It is open." % current_object.description
						
				elif current_object.opened == False:
					print "You see %s. It is closed." % current_object.description
					
		else:
			print "You see %s." % current_object.description
	
	# self prints the current inventory, your name, weapons and mood	
	elif object == 'self':
		print player_character.self_description()
			
	else:
		print "Do you see any of those around here? Locator 1: examine."
		
# take a given object
def take(object):
	current_object = current_room_objects[object]
	inv = player_character.inventory	
	
	if isinstance(current_object, Consumable) and current_object.takeable == True: 
		if current_object.consumed == False:
			print "You take the %s." % object
			current_object.takeable = False
			del current_room_objects[object]
			inv[object] = current_object
			current_object.location = 'inventory'
				
		else: 
			print "There is nothing left but %s." % current_object.consumed_description
			
	elif isinstance(current_object, Interactable) and current_object.takeable == True:
		print "You take the %s." % object
		current_object.takeable = False
		inv[object] = current_object
		# if current_object.location == 'chest':
			# print current_room_objects['chest'].contents['contents'][object]
		current_object.location = 'inventory'
		
	else:
		print "You can't take that."
		print "Locator 1: take."

# attack a given object
def attack(object):

	target = current_room_objects[object]
	
	print "You attack the %s..." % target.name
	
	if target.name == 'orc':
	
		while target.hp > 0 and player_character.player_hp > 0:
			time.sleep(0.5)
			
			if player_character.level >= target.level:
				player_attack_strength = int(player_character.str) * int(player_character.level)
				target_attack_strength = int(target.str) * int(target.level)
				
				player_defense = int(player_character.dex) * int(player_character.level)
				target_defense = int(target.dex) * int(target.level)
				
				target_damage = player_attack_strength / target_defense
				player_damage = target_attack_strength / player_defense
				
				target.hp -= target_damage
				player_character.player_hp -= player_damage
				
				print "The orc's hp is now %s." % target.hp
				print "Your hp is now %s." % player_character.player_hp

# examine the player's inventory
def inv_check():

	inv_list = []
	if is_empty(player_character.inventory) == False:
		for k, v in player_character.inventory.items():
			inv_list.append(v.description)
		x = ', '.join(inv_list)
		return x
	else:
		return "a few meagre specks of a long-since-eaten bread roll"


############################################################################################
# Meta Actions
############################################################################################

# check if a given structure is empty or not - returns true if empty
def is_empty(any_structure):
	for x in any_structure:
		return False
	else:
		return True

# break down player input into actions and objects
def action_check(player_input):
	global player_action
	global player_object
		
	player_object = None
	player_action = None
	
	player_input_list =  player_input.split()
	
	for item in player_input_list:
		if current_room_objects.has_key(item) or player_character.inventory.has_key(item) \
		or item == 'self':
				player_object = item
				# print "object in action check:", player_object
				continue
			
		elif item in valid_actions:
			player_action = item
			continue	
			
# process the player inputs and produce a response
def action_outcome(action, object):	
	
	print "\n"
	# print "object in action outcome: ", object
	
	if object in current_room_objects or object in player_character.inventory:	
		# open
		if action == 'open':
			open(object)
		# use		
		elif action == 'use':
			use(object)
		# drink / eat
		elif action == 'drink' or action ==  'eat':
			consume(action, object)
		# examine
		elif action == 'examine':
			examine(object)	
		# take
		elif action == 'take':
			take(object)
		# attack
		elif action == 'attack':
			attack(object)
		# invalid action
		elif not action == None:
			print "You can't do that with that. Locator 1: action_outcome." 
		# unclear action
		elif action == None:
			print "What do you want to do with the %s?" % object
			print "Locator 2: action_outcome."
	
	elif action == 'examine' and object == 'self':
		examine(object)	
		
	elif object not in current_room_objects and not object == None:
		print "There are none of those here. Locator 3: action_outcome."
		print action, object
	
	elif object == None:
		print "You want to do what?"
		print "Locator 4: action_outcome."
		
	player_action = None
	player_object = None

# use something on something
def use_loop(usable):
	global next_room
	
	input = get_more_info(usable)

	# if usable is a key, open the chest or door if it's currently locked
	if usable.name == 'key':
		object = current_room_objects[input]
		if object.locked == True:
			object.locked = False
			print "You unlock the %s. Locator 1: use_on_loop." % object.name
			
			if object.type == 'door':
				object.usable = True
			
		elif object.locked == False:
			print "That's not necessary. Locator 2: use_loop."
			
	# if usable is a door, open the door specified in get_more_info
	elif usable.name == 'door':
		object = current_room_objects[usable.name]
		if object.locked == True: 
			print "The door is locked"
			
		elif object.locked == False: 
			#returns new room number for the given direction
			next_room = dungeon_map[current_room.room_number][input]

# returns a specification about which object to open with a key, or which door to open in rooms
# with multiple doors (south, north, chest, door)		
def get_more_info(usable):

	while 1:
		player_input = raw_input("\n> ").lower()
		
		if player_input == '':
			continue
			
		elif player_input == 'q':
			break
			
		elif player_input in current_room_objects and usable.name == 'key':
			return player_input
			break
		
		elif player_input == 'north' and usable.name == 'door':
			print "You pass through the doorway to the north."
			return player_input
			break
			
		elif player_input == 'south' and usable.name == 'door':
			print "You pass through the doorway to the south."
			return player_input
			break
			
		else:
			print "Something went wrong in get_more_info"
			
# define current_object by checking if it's in player_character.inventory or current_room_objects
# returns valid current_object
def inv_or_room(object):
	global current_object
	
	if object in player_character.inventory:
		current_object = player_character.inventory[object]
	elif object in current_room_objects:
		current_object = current_room_objects[object]
	else:
		print "Error: inv_or_room"
		return None
		
############################################################################################
# Main Functions 
############################################################################################	
	
# do a room_check(), do action_check() on input and process an action_outcome() 
# continue on keypress Enter, and quit with 'q'
def main_loop():
	
	while 1:
		player_input = raw_input("\nWhat do you want to do?\n> ").lower()
		if player_input == '':
			continue
		elif player_input == 'q':
			break
		elif player_input == 'inv':
			print "\nYour inventory contains: %s." % inv_check()
		else:
			action_check(player_input)
			action_outcome(player_action, player_object)

# simplistic room change - non-modular for now, not dynamically altering room_counter
def room_change():
	global current_room
	global current_room_objects
		
	# store the current room's objects in room_object_library[room_number] and clear 
	# current_room_objects
	room_object_library[current_room.room_number] = copy.deepcopy(current_room_objects)
	current_room_objects.clear()
	
	# fresh populate or copy from backup stored in room_object_library
	if next_room.modified == False:
		populate_room(next_room)
		
	elif next_room.modified == True:
		current_room_objects = copy.deepcopy(room_object_library[next_room.room_number])
	
	# After leaving a room for the first time, it should be classed as modified, and the current
	# room updates
	current_room.modified = True
	
	current_room = next_room
	
	print current_room.description
	
# add to current_room_objects dict based on the consumables, monsters and interactables 
# found in the current room. Append room_numbers for the current_room to these instances.
def populate_room(current_room):
	'''Draw from the list of objects specified for the given room, creating instances of each
	object with the attributes defined in the specified dict'''
	
	# create instances
	for consumable in current_room.consumables:
		current_room_objects.update({consumable: Consumable\
		(**global_dict['consumables_dict'][consumable])})
		current_room_objects[consumable].room_number = current_room
		
	for monster in current_room.monsters:
		current_room_objects.update({monster: Monster\
		(**global_dict['monsters_dict'][monster])})
		current_room_objects[monster].room_number = current_room
	
	# problem with chest contents. cannot rely on fact that key will be generated before chest
	# produces key error for dictionary, as key may not exist
	for interactable in current_room.interactables:
		current_room_objects.update({interactable: Interactable\
		(**global_dict['interactables_dict'][interactable])})
		current_room_objects[interactable].room_number = current_room
							
############################################################################################
# Room Information & Instantiation
############################################################################################

# room_A1 Attributes
room_A1_dict = {
'description' : 
"You are standing in a small room. \
\nThere is a table in the centre, replete with a pie and a glass of milk. \
\nAn orc stands in the corner, menacing at you. \
\nBeside the orc is a hastily-made bed. \
\nYou also see a large wooden door to the north. \
\nThere is a sign on the wall, and a chest by the bed.", \
'consumables' : ['pie', 'milk'],
'interactables' : ['room', 'table', 'bed', 'door', 'sign', 'chest', 'key'], 
'monsters' : ['orc'],
'room_number' : 'room_A1',
'modified' : False,
}

# room_A2 Attributes
room_A2_dict = {
'description' : 
"You find yourself in a dark, crowded little room. \
\nThere are doors leading to the north and south. \
\nMany small statues surround you, and are placed throughout the room. \
\nA note has been pinned to one of them. \
\nYou hear the sound of dripping water.", \
'consumables' : [],
'interactables' : ['room', 'note', 'statue', 'door'], 
'monsters' : [],
'room_number' : 'room_A2',
'modified' : False,
}

# room_A3 Attributes
room_A3_dict = {
'description' : 
"A new room. \
\nThere is nothing but the door to the south and a small cake on the floor.", 
'consumables' : ['cake'],
'interactables' : ['room', 'door'], 
'monsters' : [],
'room_number' : 'room_A3',
'modified' : False,
}

############################################################################################
# Main Process
############################################################################################

current_room_objects = {}

room_object_library = {}

valid_actions = ['eat', 'drink', 'examine', 'use', 'open', 'take', 'attack']

pull_all_objects()


# create a Dungeon_Room instance for each room, and give it the attributes contained in the 
# given dictionary
room_A1 = Dungeon_Room(**room_A1_dict)

room_A2 = Dungeon_Room(**room_A2_dict)

room_A3 = Dungeon_Room(**room_A3_dict)

# specify linked rooms in the dungeon, for room changes
dungeon_map = {
'room_A1' : {'north' : room_A2},
'room_A2' : {'north' : room_A3, 'south' : room_A1},
'room_A3' : {'south' : room_A2},
}

# set starting room
current_room = room_A1

# initiate the global variable next_room
next_room = room_A2

# populate the room with usable objects

populate_room(current_room)

# comment/uncomment to allow user to personalise the character
# personalisation()

# change to player_name if personalisation() active
player_character = Player(**global_dict['player_info_dict']['Andrew'])

print "\n"

print current_room.description

print player_character.self_description()



main_loop()
