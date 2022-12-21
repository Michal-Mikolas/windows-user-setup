import os, sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from matt.matt import Matt

class Settings:
	def __init__(self, cache_path:str) -> None:
		self.matt = Matt(cache_file=cache_path)
		self.matt.set_ui({
			'run_title': ['img/run_title_w10.png'],
		})
