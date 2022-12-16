

"""
https://winaero.com/ms-settings-commands-in-windows-10/
http://woshub.com/ms-settings-uri-commands-windows-11/
"""
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

