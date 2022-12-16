
class Windows():
	def __init__(self, cache_path:str) -> None:
		self.cache_path = cache_path

	def run(self, cmd:str):
		pass

	def settings_language(self):
		# W10: ms-settings:regionlanguage-setdisplaylanguage
		# W11: ms-settings:regionlanguage                     # works on W10 as well
		self.run('ms-settings:regionlanguage')

	def search(self, name:str):
		pass

