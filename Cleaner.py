import re

raw_data = open('rawSerials.txt', 'r')
working_copy = raw_data.read()
raw_data.close()

# Replace 'O' (oh) with '0' (zero).
working_copy = re.sub(r'[oO]', '0', working_copy)

# Replace 'W' with 'VV'.
working_copy = re.sub(r'[wW]', 'VV', working_copy)

# Remove special characters.
working_copy = re.sub(r'[^ \t\n\r\f\va-zA-Z0-9]*', '', working_copy)

# Remove extra white space.
working_copy = re.sub(r'[ \t\n\r\f\v]+', '', working_copy)

if working_copy[0] == 'S' or working_copy[0] == 'C':
    i = 14 # Modem
    sn_and_space = 15
elif working_copy[0] == 'A':
    i = 11 # WAP
    sn_and_space = 12
else:
    i = 9 # STB or PVR
    sn_and_space = 10

# Re-add white space between SNs.
while i < len(working_copy):
        working_copy = working_copy[:i] + " " + working_copy[i:]
        i = i+sn_and_space

output = open("cleanSNs.txt", "w")
output.write(working_copy)
output.close()
