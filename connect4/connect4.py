# -*- coding: utf-8 -*-
from random import randint

X=0
Y=1
VERTICAL = 0
HORIZONTAL = 1
DIAGONAL_R = 2
DIAGONAL_L = 3
dim = (7,6)


'''
TODO:
	[X] Implement basic game logic
	[X] Setup project GIT repository
	[] Implement play system (ie players, game loop)
	[] Implement state node class
	[] Implement random AI player
		[] AI can expand current state to all possible children nodes
		[] AI picks one child node at random
	[] Implement random state expansion to level n
		- Returns a state? Node?.. Board?
	[] Implement intelligent AI player
		-Iterative deepening Alpha-Beta
		[] Investigate Alpha-Beta optimizations (Negamax/scout,killer heu..)
		[] Define state evaluation function
		[] Implement state evaluation function
		[] Implement basic Alpha-Beta/Minimax to depth n
		[] Improve with optimizations
			[] Transposition tables
			[] Iterative deepening
			[] Timed iterative deepening
			
SECONDARY:
	[] Optimize getWinner() search
		- Stop searching a particular group once 4 in a row is impossible
		- 'if count + ((dim[ceil(?)]-1)-i) < 4: break'

BUGS:
	No known bugs
'''

class Board:
	def __init__( self, turns=0, state=None ):
		self.state = state or getEmpty()
		self.turn = -1
		self.playRandom(turns)
		
	def insert( self, col ):
		i = dim[Y]-1
		while i >= 0:
			if self.state[i][col] == 0 and i > 0:
				i -= 1
				
			elif i < dim[Y]-1:
				if self.state[i][col]: self.state[i+1][col] = self.turn
				else: self.state[i][col] = self.turn
				self.turn = -self.turn
				return True
			
			else: return False
		return False
		
	def reset( self ):
		self.state = getEmpty()
		
	def playRandom( self, turns=10 ):
		for t in range( turns ):
			self.insert( randint(0,dim[X]-1) )
			
	def winner( self ):
		return getWinner( self.state )
	
	def printOut( self ):
		printBoard( self.state )

def getEmpty( x=dim[X], y=dim[Y] ):
	return [[0 for i in range(x)] for j in range(y)]

def printBoard( state ):
	for j in range(len(state)):
		for i in range(len(state[j])):
			if state[(dim[Y]-1)-j][i] == 1: print 'X',
			elif state[(dim[Y]-1)-j][i] == -1: print 'O',
			else: print '.',
		print ''
	print ''

def isGameEnd( state ):
	for y in range(len(state)):
		for x in range(len(state[y])):
			if state[y][x] == 0: return False
	return True
	
def getStart( pos, d ):
	start = [0,0]
	
	if d == VERTICAL or d == HORIZONTAL:
		ceil = 1-d
		start[1-ceil] = pos[1-ceil]

	elif d == DIAGONAL_R:
		ceil = min([(dim[i]-pos[i],i) for i in range(len(dim))])[1]
		start[ceil] = dim[ceil] - (dim[ceil] - pos[ceil] + pos[1-ceil])

	elif d == DIAGONAL_L:
		ceil = min( (dim[Y]-pos[Y],1), (pos[X],0) )[1]
		start[1-ceil] = ceil*(dim[X] - 1)
		if ceil: start[ceil] = pos[ceil] - (dim[1-ceil] - 1) - pos[1-ceil]
		else: start[ceil] = dim[ceil] - pos[ceil] + pos[1-ceil] - 1
		
	return start
	
def getMidPoint( direction, increment ):
	if direction == VERTICAL: return (increment,dim[Y]/2)
	else: return (dim[X]/2,increment)

def isLegal( pos ):
	for i in range(len(pos)):
		if pos[i] < 0 or pos[i] >= dim[i]: return False
	
	return True
	
def getWinner( state ):
	dmap = [(0,1),(1,0),(1,1),(-1,1)]
	toCheck = state
	
	for direction in range(4):
		print ['VERTICAL','HORIZONTAL','DIAG_R','DIAG_L'][direction]
		for i in range(dim[min(direction,1)]):
			midpoint = getMidPoint( direction, i )
			toCheck = state[midpoint[Y]][midpoint[X]]
			print midpoint, toCheck
			
			
			if toCheck:
				pos = getStart( midpoint, direction )
				count = 0
				
				while isLegal( pos ):
					if state[pos[Y]][pos[X]] == toCheck: count += 1
					else: count = 0
					
					print '    %(pos)s:\tv[%(v)d]\tc[%(m)c,%(c)d]' % {'pos':str(pos),'v':state[pos[Y]][pos[X]], 'm':{-1:'O',1:'X'}[toCheck], 'c':count}
					
					if count == 4: return toCheck		
					pos[X] += dmap[direction][X]
					pos[Y] += dmap[direction][Y]
					
	return 0
