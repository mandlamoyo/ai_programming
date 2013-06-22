# -*- coding: utf-8 -*-
from random import choice, shuffle
from copy import deepcopy

class Node:
	def __init__( self, state, turn ):
		self.state = state
		self.turn = turn
		self.score = getStateScore(self.state,self.turn) #0
		self.depth = nodeDepth( self.state )
		self.alphabeta = [99,-99] #[beta,alpha]
		self.best = None

	def expand( self ):
		return [Node(s,-self.turn) for s in fullExpand(self.state,self.turn)]
		
	def getID( self ):
                return getStateID( self.state )

class Player:
	def __init__( self ):
		self.player = 0
		
	def play( self, state ):
		pass
		
class RandomPlayer( Player ):
	def play( self, state ):
		return choice( fullExpand( state, self.player ))
		
class CleverPlayer( Player ):
	def play( self, state ):
		return getBestMove( Node(state, self.player) ).state
		
class HumanPlayer( Player ):
	def play( self, state ):
		printBoard( state )
		print 'Enter move coordinates:'
		x_pos = int(raw_input( 'x coordinate: ' ))
		y_pos = int(raw_input( 'y coordinate: ' ))
		print ''

		ns = state[:]
		ns[y_pos][x_pos] = self.player
		return ns

INVESTIGATED = 0
EXPANDED = 1
TT_INS = 2
TT_HIT = 3
TT_GET = 4
#BEST,DEPTH,SCORE
BEST = 0
DEPTH = 1
SCORE = 2
MIN = 0
MAX = 1
ALPHA = 1
BETA = 0
WIN = 10
BLOCK = 5
SCOREPOSITIONS = [[[0,0],[1,1],[2,2]],[[0,2],[1,1],[2,0]],[[0,0],[0,1],[0,2]],[[1,0],[1,1],[1,2]],[[2,0],[2,1],[2,2]],[[0,0],[1,0],[2,0]],[[0,1],[1,1],[2,1]],[[0,2],[1,2],[2,2]]]
getPlayer = lambda x : x/abs(x)
tt = {}


def encode(s):
	ns = [[ s[0][0], s[0][2], s[2][2], s[2][0] ], [ s[0][1], s[1][2], s[2][1], s[1][0] ],[ s[1][1] ]]
	for i in range(len(ns)):
		for j in range(len(ns[i])):
			if ns[i][j] == -1: ns[i][j] = 2
	return ns

def decode(s):
	ns = [[s[0][0],s[1][0],s[0][1]],[s[1][3],s[2][0],s[1][1]],[s[0][3],s[1][2],s[0][2]]]
	for i in range(len(ns)):
		for j in range(len(ns[i])):
			if ns[i][j] == 2: ns[i][j] = -1
	return ns

def mirror(os):
	s = deepcopy(os)
	s[0][0], s[0][3] = s[0][3], s[0][0]
	s[0][1], s[0][2] = s[0][2], s[0][1]
	s[1][0], s[1][2] = s[1][2], s[1][0]
	return s
        
def shift( s ):
	if len(set(s[0])) > 1: c = 0
	elif len(set(s[1])) > 1: c = 1
	else: return s

	while s[c][0] != min( s[c] ) or s[c][0] == s[c][-1]:
		for i in range(2): s[i] = s[i][1:] + [s[i][0]]
	return s

def getStateID(state):
	estate = encode(state)
	code = [shift( estate ),shift( mirror( estate ))]
        if not code[0] and not code[1]:
                print code
                print state
                
	code.sort()
	return hash( tuple( sum( code[0]+code[1], [] )))
        

def getRandom(d=5,slim=5):
        acceptible = False
        while not acceptible:
                n = [1,11]
                while n[1] >= slim: n = explore(d)
                print n[1]
                printBoard(n[0].state)
                acceptible = {'y':True,'n':False}[raw_input('acceptible board?: ')]
                print ''
        return n

def alphaBetaTT(node):
        global tt
        tt = {}
        res = alphaBetaVerbose(node,t_on=True)
        return res, len(tt)
        
def alphaBetaVerbose( node, tdepth=10, ind=0, t_on=False ):
	spac = '     '
        stats = [0,0,0,0,0] #INVESTIGATED,EXPANDED,TT_INS,TT_HIT,TT_GET
        global tt
        
	#print spac * ind,
	if not isGameEnd(node.state) and node.depth < tdepth and not isWinner(node.state):
		nType = {-1:MIN,1:MAX}[node.turn]
		
		#print {MIN:'MIN ',MAX:'MAX '}[nType] + 'node ' + str(node.depth),
		children = node.expand()
		shuffle(children)
		
		#print '[' + str(len(children)) + '] ('
		
                if t_on:
                        nid = node.getID()
                        if nid in tt:
                                #print spac * ind,
                                stats[TT_HIT] += 1
                                res = tt[nid]
                                if res[DEPTH] >= node.depth:
                                        #print 'Node already explored at depth (%(d)d), with score (%(s)d)' % {'d':res[DEPTH],'s':res[SCORE]}
                                        node.best = Node( res[BEST], -node.turn )
                                        stats[TT_GET] += 1
                                        return res[SCORE], stats
                                #else: 
                                        #print 'Node previously encountered..'
                                        #print spac * ind,
                                        
		for child in children:
                        stats[EXPANDED] += 1
			if node.alphabeta[ALPHA] >= node.alphabeta[BETA]:
				#print spac * ind,
				#print str(node.alphabeta[ALPHA]) + '!' + str(node.alphabeta[BETA]) + ': ' + str(node.alphabeta[nType])
				return node.alphabeta[nType], stats
			
                        stats[INVESTIGATED] += 1
                        
			child.alphabeta[nType] = node.alphabeta[nType]
			res = alphaBetaVerbose(child,tdepth,ind+1,t_on)
			if (res[0] < node.alphabeta[nType] and nType == MIN) or (res[0] > node.alphabeta[nType] and nType == MAX):
				node.alphabeta[nType] = res[0]
				node.best = child
			#print spac * ind,
			#print ' [a:' + str(node.alphabeta[ALPHA]) + ', b:' + str(node.alphabeta[BETA]) + ']'
                        stats = [stats[i] + res[1][i] for i in range(len(stats))]

                        
                if t_on and nid not in tt:
                        tt[nid] = [node.best.state, node.depth, node.alphabeta[nType]] 
                        stats[TT_INS] += 1
                        ##print spac*ind + 'Node inserted into tt'
                        
		#print spac * ind,
		#print '): ' + str(node.alphabeta[nType]) + '\n'
		return node.alphabeta[nType], stats
	
	else:
		#print 'LEAF ' + str(node.depth) + ': ' + str(node.score)
		#print spac * ind,
		#print '--',
		#print node.depth < tdepth, isGameEnd(node.state), isWinner(node.state)
		#print ''
		return node.score, stats

def getBestMove( node, tdepth=10, comments=True ):
	res = alphaBeta( node, tdepth )
	if not comments: return node.best
	
	fn = node
	mlist = []
	
	while fn.best != None:
		mlist.append(fn)
		fn = fn.best
	mlist.append(fn)
	
	for n in mlist:
		printBoard(n.state)
		print '------\n'
		
	print '\n'
	return node.best
	
def alphaBeta( node, tdepth, start=True, t_on=True ):
        global tt
        if start:
                tt = {}
                start = False
                
	if not isGameEnd(node.state) and node.depth < tdepth and not isWinner(node.state):
		nType = {-1:MIN,1:MAX}[node.turn]
		
                
                if t_on:
                        nid = node.getID()
                        if nid in tt:
                                res = tt[nid]
                                if res[DEPTH] >= node.depth:
                                        node.best = Node( res[BEST], -node.turn )
                                        return res[SCORE]
                                        
		for child in node.expand():
			if node.alphabeta[ALPHA] >= node.alphabeta[BETA]:
				return node.alphabeta[nType]
			
			child.alphabeta[nType] = node.alphabeta[nType]
			res = alphaBeta(child,tdepth)
			if (res < node.alphabeta[nType] and nType == MIN) or (res > node.alphabeta[nType] and nType == MAX):
				node.alphabeta[nType] = res
				node.best = child
				
                if t_on and nid not in tt: tt[nid] = [node.best.state, node.depth, node.alphabeta[nType]]
		return node.alphabeta[nType]
	
	else:
		return node.score

def nodeDepth( state ):
	depth = 1
	for i in range(len(state)):
		for j in range(len(state[i])):
			if state[i][j] != 0:
				depth += 1
	return depth


def explore( d ):
	start = Node(emptyBoard(),1)
	current = start
	while( current.depth < d and not isGameEnd( current.state ) and not isWinner( current.state )):
		current = choice( current.expand() )

	return [current,getStateScore(current.state,current.turn)]


def playGame( players, board=None ):
	if not board: board = emptyBoard()
	for p in range(len(players)): players[p].player = [-1,1][p]

	while( not isGameEnd( board )):
		for p in players:
			board = p.play( board )
			#print board
                        print alphaBetaVerbose( Node(board,-p.player) )
                        print alphaBetaTT( Node(board,-p.player) )
                        print ''
                        
			res = isWinner(board)
			if res: return res
	return 0
	

def getStateScore( state, pTurn ):
	
	winnable = [set(),set()]
	blockable = 0
	desirability = 0
	
	for group in SCOREPOSITIONS:
		res = getGroupScore( group, state )
		#print group, res

		if abs(res) == 3: return WIN * getPlayer(res)
		if abs(res) == 2:
			if getPlayer(res) == pTurn: return WIN * getPlayer(res)
			winnable[{-1:0,1:1}[getPlayer(res)]].add(tuple(getDependent(group,state)))

		else: desirability += res

	#print winnable, desirability
	for w in range(len(winnable)):
		#print winnable[w]
		if len(winnable[w]) > 1: return WIN * [-1,1][w]
		if len(winnable[w]) == 1:
			blockable = [-1,1][w]
			
	#print blockable, desirability
	return desirability + blockable * BLOCK


def getGroupScore( group, state ):
	return sum([ state[coord[0]][coord[1]] for coord in group ])

def getDependent( group, state ):
	return [coord for coord in group if state[coord[0]][coord[1]] == 0][0]

def isWinner( state ):
	for group in SCOREPOSITIONS:
		res = getGroupScore( group, state )
		if abs(res) == 3: return getPlayer(res)
	return 0
		
def isGameEnd( state ):
	for y in range(len(state)):
		for x in range(len(state[y])):
			if state[y][x] == 0: return False
	return True
	
	
def expand( state, t, n=-1 ):
	if n == -1:
		return [ expand(state[:],t,i) for i in range(len(state)) if state[i] == 0 ]
	else:
		state[n] = t
		return state
	
def fullExpand( state, t ):
	slist = []
	for g in range(len(state)):
		for expandedGroup in expand(state[g],t):
			ns = state[:]
			ns[g] = expandedGroup
			slist.append(ns)
			
	return slist

def emptyBoard():
	return [[0 for i in range(3)] for j in range(3)]

def printBoard( state ):
	for j in range(len(state)):
		for i in range(len(state[j])):
			if state[j][i] == 1: print 'X',
			elif state[j][i] == -1: print 'O',
			else: print ' ',
		print ''
	print''

if __name__ == '__main__':
	run = raw_input('Play new game? (y/n): ').lower() == 'y'
	while run:
		players = [HumanPlayer(),CleverPlayer()]
		if raw_input('Would you like to go first? (y/n): ').lower() != 'y':
			players.reverse()
			
		winner = playGame( players )
		print {-1:"Noughts win!",1:"Crosses win!",0:"It's a draw!"}[winner]
		run = raw_input('Play again? (y/n): ').lower() == 'y'