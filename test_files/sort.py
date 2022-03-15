from natsort import natsorted # pip install natsort

# Example list of strings
a = ['1', '10', '2', '3', '11']

# Your array may contain strings
y = ['string11', 'string3', 'string1', 'string10', 'string100']
x = natsorted(y)
print(x)