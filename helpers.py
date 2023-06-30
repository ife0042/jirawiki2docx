import webcolors

def color_name_to_hex(color_name):
	try:
		# Get the RGB values for the given color name
		rgb = webcolors.name_to_rgb(color_name)
		# Convert the RGB values to a hex code
		hex_code = "#{:02x}{:02x}{:02x}".format(rgb.red, rgb.green, rgb.blue)
		return hex_code
	except ValueError:
		# Handle the case when an invalid color name is provided
		print("Invalid color name.")
		return None