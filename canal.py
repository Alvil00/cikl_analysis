import sys
import weakref
import openpyxl
from openpyxl.styles.borders import Border, Side
from openpyxl.styles.alignment import Alignment
from openpyxl.styles.fonts import Font
import collections
import os
import re
import enum
import argparse

def parse_args(arguments):
	p = argparse.ArgumentParser('This is script to read an collect data from cycle vtu calculation')
	pg = p.add_mutually_exclusive_group()
	pg.add_argument('-n', type=int, default=10, help='number of extracted records')
	pg.add_argument('-l', type=int, nargs='+', help='specified list of node numbers')
	p.add_argument('--limit', type=float, default=1E-8, help='set the limit of extracted types of cycle')
	p.add_argument('--outfile', type=str, default='table.xlsx', help='name of output file, default = table.xlsx')
	r = p.parse_args(arguments)
	return r

class CycleTypeRecord():
	
	def __init__(self, parent, first_id, second_id, saf ,sfmax, sfmin, tmax, tmin, r, ndop, n, a):
		self._first_id = first_id
		self._second_id = second_id
		self._saf = saf
		self._sfmax = sfmax
		self._sfmin = sfmin
		self._tmax = tmax
		self._tmin = tmin
		self._r = r
		self._ndop = ndop
		self._n = n
		self._a = a
		if parent is None:
			self._parent = None
		elif isinstance(parent, CycleTypeTable):
			self._parent=weakref.ref(parent)
			self._parent().append(self)
		else:
			raise ValueError
	
	@property
	def parent(self):
		if self._parent is None:
			return None
		else:
			return self._parent()
			
	
	@property
	def first_id(self):
		return self._first_id
	
	@property
	def second_id(self):
		return self._second_id
		
	@property
	def saf(self):
		return self._saf
	
	@property
	def sfmax(self):
		return self._sfmax
	
	@property
	def sfmin(self):
		return self._sfmin
		
	@property
	def tmax(self):
		return self._tmax
	
	@property
	def tmin(self):
		return self._tmin
		
	@property
	def r(self):
		return self._r
	
	@property
	def ndop(self):
		return self._ndop
	
	@property
	def n(self):
		return self._n
		
	@property
	def a(self):
		return self._a
	


class CycleTypeTable(collections.UserList):
		
	def __init__(self, nodenum, parent=None):
		self._nodenum = nodenum
		super().__init__()
		if parent is None:
			self._parent = None
		elif isinstance(parent, CycleTypeManagerTable):
			self._parent=weakref.ref(parent)
			self._parent()[nodenum] = self
		else:
			raise ValueError
	
	@property
	def nodenum(self):
		return self._nodenum
	
	@property
	def parent(self):
		if self._parent is None:
			return None
		else:
			return self._parent()
	
	def print_table(self):
		print("fid   sid  sfmax   sfmin   saf      tmin   tmax   r      ndop       n        a")
		for i in self:
			if i.a > 1E-8:
				print("{fid:>3} - {sid:<3} {sfmax:>8.2f}{sfmin:>8.2f}{saf:>8.2f}{tmin:>7.1f}{tmax:>7.1f}{r:>7.2f}{ndop:>10.0f}{n:>10.1f} {a:>11.4e}".format(fid=i.first_id, sid=i.second_id, saf=i.saf, n=i.n, a=i.a, sfmax=i.sfmax, sfmin=i.sfmin, tmin=i.tmin, tmax=i.tmax, r=i.r, ndop=i.ndop))

class CycleTypeManagerTable(collections.UserDict):
	class AccumulatedFatigueDamageFileLineContext(enum.Enum):
		FIND_NODE_NUM = 0
		FIND_NUM_COMPONENT_NUM = 1
		FIND_NODE_BM = 2
		READ_RECORD = 3
		
	def __init__(self, node_table, local_reduced_stress_manager_table):
		super().__init__()
		if node_table is None:
			self._node_table = None
		elif isinstance(node_table, NodeTable):
			self._node_table=weakref.ref(node_table)
		else:
			raise ValueError
	
		if local_reduced_stress_manager_table is None:
			self._local_reduced_stress_manager_table = None
		elif isinstance(local_reduced_stress_manager_table, LocalReducedStressManagerTable):
			self._local_reduced_stress_manager_table=weakref.ref(local_reduced_stress_manager_table)
		else:
			raise ValueError
		
	@property
	def node_table(self):
		if self._node_table is None:
			return None
		else:
			return self._node_table()
			
	@property
	def local_reduced_stress_manager_table(self):
		if self._local_reduced_stress_manager_table is None:
			return None
		else:
			return self._local_reduced_stress_manager_table()
			
	
	def parse_accumulated_fatigue_damage_file(self, file):
		current_context = self.AccumulatedFatigueDamageFileLineContext.FIND_NODE_NUM
		current_table = None
		current_nodenum = None
		nessesery_component = None
		necessery_base_moment = None
		find_node_num_pattern = re.compile('(?<=\>\sCalculation\snode:\s)\d+')
		find_node_component_pattern = re.compile('(?<=\>\sComponent\snumber:\s)\d+')
		find_node_basemoment_pattern = re.compile('(?<=\>\sBase\scalculated\smoment\sof\stime\s)\d+')
		with open(file, mode='r') as f:
			for line in f:
				if current_context == self.AccumulatedFatigueDamageFileLineContext.FIND_NODE_NUM:
					result = re.search(find_node_num_pattern, line)
					if result:
						current_nodenum = int(result.group(0))
						if self.node_table.get(current_nodenum):
							nessesery_component = self.node_table[current_nodenum].component
							necessery_base_moment = self.node_table[current_nodenum].base_moment
							current_table = CycleTypeTable(int(result.group(0)), self)
							current_context = self.AccumulatedFatigueDamageFileLineContext.FIND_NUM_COMPONENT_NUM
				elif current_context == self.AccumulatedFatigueDamageFileLineContext.FIND_NUM_COMPONENT_NUM:
					result = re.search(find_node_component_pattern, line)
					if result and nessesery_component == int(result.group(0)):
						current_context = self.AccumulatedFatigueDamageFileLineContext.FIND_NODE_BM
				elif current_context == self.AccumulatedFatigueDamageFileLineContext.FIND_NODE_BM:
					result = re.search(find_node_basemoment_pattern, line)
					if result and necessery_base_moment == int(result.group(0)):
						current_context = self.AccumulatedFatigueDamageFileLineContext.READ_RECORD
						header = 2
				elif current_context == self.AccumulatedFatigueDamageFileLineContext.READ_RECORD:
					if header > 0:
						header -= 1
					else:
						temp_list = line.replace(',','.').strip().split()
						try:
							CycleTypeRecord(current_table, int(temp_list[0]), int(temp_list[2]), float(temp_list[7]), float(temp_list[5]), float(temp_list[6]), float(temp_list[9]), float(temp_list[8]), float(temp_list[10]), float(temp_list[19]), float(temp_list[20]), float(temp_list[21]))
						except (ValueError, IndexError ) as vi:
							current_table.sort(key = lambda a: a.a, reverse=True)
							current_context = self.AccumulatedFatigueDamageFileLineContext.FIND_NODE_NUM
	
class LocalReducedStressRecord():
	
	def __init__(self, num, parent= None,temp:float=20.0, si:float=0.0, sj:float=0.0, sk:float=0.0):
		self._num = num
		self._temp = temp
		self._list = [si, sj, sk, None, None, None]
		if parent is None:
			self._parent = None
		elif isinstance(parent, LocalReducedStressTable):
			self._parent=weakref.ref(parent)
			self._parent()[num] = self
		else:
			raise ValueError
	
	@property
	def parent(self):
		if self._parent is None:
			return None
		else:
			return self._parent()
	
	@property
	def num(self):
		return self._num
		
	
	@property
	def temp(self):
		return self._temp
		
	@property
	def si(self):
		return self._list[0]
	
	@property
	def sj(self):
		return self._list[1]
		
	@property
	def sk(self):
		return self._list[2]
		
	@property
	def sij(self):
		if self._list[3] is None:
			self._list[3] = self._list[0] - self._list[1]
		return self._list[3]
	
	@property
	def sjk(self):
		if self._list[4] is None:
			self._list[4] = self._list[1] - self._list[2]
		return self._list[4]
		
	@property
	def sik(self):
		if self._list[5] is None:
			self._list[5] = self._list[0] - self._list[2]
		return self._list[5]
	
	@property
	def vec(self):
		return self._list.copy()

class LocalReducedStressTable(collections.UserDict):
	def __init__(self, nodenum, parent):
		self._nodenum = nodenum
		super().__init__()
		if parent is None:
			self._parent = None
		elif isinstance(parent, LocalReducedStressManagerTable):
			self._parent=weakref.ref(parent)
			self._parent()[nodenum] = self
		else:
			raise ValueError

	
	@property
	def nodenum(self):
		return self._nodenum
	
	@property
	def parent(self):
		if self._parent is None:
			return None
		else:
			return self._parent()

	def print_table(self):
		print('id        temp      sij       sjk       sik')
		for moment, m in self.items():
			print("{moment:<10}{temp:<10.1f}{sij:<10.2f}{sjk:<10.2f}{sik:<10.2f}".format(moment=moment, temp=m.temp, sij=m.sij, sjk=m.sjk, sik=m.sik))

			

class LocalReducedStressManagerTable(collections.UserDict):
	class LocalReducedStresFileLineContext(enum.Enum):
		FIND_NODE_NUM = 0
		FIND_NODE_BM = 1
		READ_RECORD = 2
		
	def __init__(self, node_table):
		super().__init__()
		if node_table is None:
			self._node_table = None
		elif isinstance(node_table, NodeTable):
			self._node_table=weakref.ref(node_table)
		else:
			raise ValueError
	
	@property
	def node_table(self):
		if self._node_table is None:
			return None
		else:
			return self._node_table()
	
	def parse_local_redused_stress_file(self, file, verbose=False):
		current_context = self.LocalReducedStresFileLineContext.FIND_NODE_NUM
		current_table = None
		current_nodenum = None
		necessery_base_moment = None
		find_node_num_pattern = re.compile('(?<=\>\sCalculation\snode\s)\d+')
		find_node_basemoment_pattern = re.compile('(?<=\>\>moment\s)\d+(?=\s-\>\scalculation\sresults\sTable)')
		start_header_passer = False
		with open(file, mode='r') as f:
			for line in f:
				if current_context == self.LocalReducedStresFileLineContext.FIND_NODE_BM:
					result = re.search(find_node_basemoment_pattern, line)
					if result and necessery_base_moment == int(result.group(0)):
						if verbose:
							print('Find necessery_base_moment = {}'.format(necessery_base_moment))
						current_context = self.LocalReducedStresFileLineContext.READ_RECORD
						start_header_passer = True
				elif current_context == self.LocalReducedStresFileLineContext.READ_RECORD:
					if not start_header_passer:
						try:
							temp_list = line.replace(',','.').strip().split()
							LocalReducedStressRecord(int(temp_list[0]), current_table, float(temp_list[1]), float(temp_list[5]), float(temp_list[6]), float(temp_list[7]))
						except (ValueError, IndexError):
							current_context = self.LocalReducedStresFileLineContext.FIND_NODE_NUM
					else:
						start_header_passer = False
				if current_context == self.LocalReducedStresFileLineContext.FIND_NODE_NUM:
					start_header_passer = False
					result = re.search(find_node_num_pattern, line)
					if result:
						current_nodenum = int(result.group(0))
						if self.node_table.get(current_nodenum):
							if verbose:
								print('Find node header = {}'.format(current_nodenum))
							necessery_base_moment = self.node_table[current_nodenum].base_moment
							current_table = LocalReducedStressTable(int(result.group(0)), self)
							current_context = self.LocalReducedStresFileLineContext.FIND_NODE_BM
						
class NodeRecord():
	def __init__(self, num, parent=None, damage:float=0.0, base_moment:int=0, component:int=0):
		self._num = num
		self._damage = damage
		self._base_moment = base_moment
		self._component = component
		if parent is None:
			self._parent = None
		elif isinstance(parent, NodeTable):
			self._parent=weakref.ref(parent)
			self._parent()[num] = self
		else:
			raise ValueError

			
	@property
	def parent(self):
		if self._parent is None:
			return None
		else:
			return self._parent()
	
	@property
	def num(self):
		return self._num
	
	@property
	def damage(self):
		return self._damage
	
	@property
	def base_moment(self):
		return self._base_moment
	
	@property
	def component(self):	
		return self._component
	
	@damage.setter
	def damage(self, value:float):
		if self._damage == 0.0 and value >= 0.0:
			self._damage = value
		else:
			raise ValueError
		
	@base_moment.setter
	def base_moment(self, value:int):
		if self._base_moment == 0 and value >= 0:
			self._base_moment = value
		else:
			raise ValueError(value,self._base_moment)
			
	@component.setter
	def component(self, value:int):
		if self._component == 0 and value in (1, 2, 3):
			self._component = value
		else:
			raise ValueError
		
class NodeTable(collections.UserDict):
	
	def __init__(self):
		super().__init__()
		self._dindex = None
		
	def parse_base_moments(self, file):
		extract = lambda typ, line, sign: typ(line.split(sign)[1].strip())
		with open(file, mode='r') as f:
			for line in f: 
				if line.startswith('Calculation'):
					current_node = extract(int, line, ':')
					NodeRecord(current_node, self)
				elif line.startswith('a = '):
					self[current_node].damage = extract(float, line.replace('.','').replace(',','.'), '=')
				elif line.startswith('Base calculated moment of time'):
					self[current_node].base_moment = extract(int, line.replace(',',''), ':')
				elif line.startswith('reduced sterss component'):
					self[current_node].component = extract(int, line.replace('.',''), ':')
				else:
					pass
	
	def get_damage_index(self, num=None):
		if self._dindex is None:
			self._dindex = sorted(self.values(), key=lambda a: a.damage,reverse=True)
		if num is None:
			return self._dindex
		else:
			return self._dindex[:num]
	
	def print_table(self, limit=0, sort_by_damage=False, ):
		print('nodenum   damage          bm    component')
		if sort_by_damage:
			ld = list(self.items())
			ld.sort(key=lambda a: a[1].damage,reverse=True)
		else:
			ld = self.items()
		if limit is None or limit <= 0:
			for nodenum, node in ld:
				print("{node:<10}{damage:<10.5e}{bm:6}{c:>6}".format(node=nodenum, damage=node.damage, bm=node.base_moment, c=node.component))
		else:
			for num, (nodenum, node) in enumerate(ld):
				if num >= limit:
					break
				print("{node:<10}{damage:<10.5e}{bm:6}{c:>6}".format(node=nodenum, damage=node.damage, bm=node.base_moment, c=node.component))
					
def save_in_workbook(manager_table, necessery_nodes=None, worksheet_name='ma', limit=1E-8):
	mb = openpyxl.Workbook()
	chwidth = (("Тип цикла", 12),
	          ("σFmax",     8.25),
	          ("σFmin",     8.25),
	          ("σaF",       8.25),
	          ("Tmin",      8.25),
	          ("Tmax",      8.25),
	          ("r",         8),
	          ("[N]",       11),
	          ("N",         9),
	          ("a",         10))
	font = Font(name='Times New Roman', size=12)
	thin_border = Border(left=Side(style='thin'),
                     right=Side(style='thin'),
                     top=Side(style='thin'),
                     bottom=Side(style='thin'))
	cent_alignment = Alignment(horizontal="center",
                           vertical="center",
                           wrap_text=True)
	if necessery_nodes is None:
		necessery_nodes = manager_table.keys()
	for node in necessery_nodes:
		is_empty = True
		sheet_name = '{}n'.format(node)
		ms = mb.create_sheet(sheet_name)
		ct = manager_table[node]
		for rownum, item in enumerate(filter(lambda a: a.a > limit, ct), 2):
			ms.cell(row=rownum, column=1).value = "{}-{}".format(item.first_id, item.second_id)
			ms.cell(row=rownum, column=2).value = item.sfmax   
			ms.cell(row=rownum, column=3).value = item.sfmin
			ms.cell(row=rownum, column=4).value = item.saf
			ms.cell(row=rownum, column=5).value = item.tmin
			ms.cell(row=rownum, column=6).value = item.tmax
			ms.cell(row=rownum, column=7).value = item.r
			ms.cell(row=rownum, column=8).value = item.ndop
			ms.cell(row=rownum, column=9).value = item.n
			ms.cell(row=rownum, column=10).value = item.a
			is_empty = False
		if not is_empty:
			for cnum, (chvalue, cwidth) in enumerate(chwidth, 1):
				hcell = ms.cell(row=1, column=cnum)
				hcell.value = chvalue
				ms.column_dimensions[hcell.column].width = cwidth
			ms.cell(row=rownum+1, column=10).value = "=SUM(J2:J{})".format(rownum)
			ms.merge_cells(start_row=rownum+1, start_column=1, end_row=rownum+1, end_column=9)
			ms.cell(row=rownum+1, column=1).value = "Итоговая накопленная усталостная поврежденность"
			for row in ms['A1:J{}'.format(rownum+1)]:
				for cell in row:
					cell.border = thin_border
					cell.alignment = cent_alignment
					cell.font = font
					if cell.column in 'BCDEFH':
						cell.number_format = '0'
					elif cell.column == 'G':
						cell.number_format = '0.00'
					elif cell.column == 'I':
						cell.number_format = '0.0'
					elif cell.column == 'J':
						cell.number_format = '0.0E+0'
	try:
		mb.save("{}".format(worksheet_name))
	except PermissionError:
		print('ERROR: Please close the excel file')
	
		
def main():
	args = parse_args(sys.argv[1:])
	if not args.l:
		nnodes = args.n
	else:
		nnodes = None
	nt = NodeTable()
	nt.parse_base_moments(list(filter(lambda a: a.startswith('BaseMoments'), os.listdir()))[0])
	if nnodes:
		nt.print_table(sort_by_damage=True, limit=nnodes)
	lmt = LocalReducedStressManagerTable(nt)
	lmt.parse_local_redused_stress_file(list(filter(lambda a: a.startswith('Report (Local Reduced Stress)'), os.listdir()))[0])
	ctt = CycleTypeManagerTable(nt, lmt)
	ctt.parse_accumulated_fatigue_damage_file(list(filter(lambda a: a.startswith('Report (Accumulated Fatigue Damage)'), os.listdir()))[0])
	if nnodes:
		nn = list(map(lambda a: a.num, nt.get_damage_index(nnodes)))
		save_in_workbook(ctt, nn, args.fname, args.limit)
	else:
		save_in_workbook(ctt, args.l, args.fname, args.limit)
	
	
if __name__ == "__main__":
	main()