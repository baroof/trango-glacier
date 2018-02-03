#!/usr/bin/python
# This program adds up integers
# expects a comma delimited string: 1,2,5,6
# returns the count of integers (4) and sum (14)
import sys

while True:
    my_s = raw_input("Gimme: ")
    my_l = my_s.split(",")
    if len(my_s) == 0:
        print("I quit!\n")
        break
    else:
        print '# of digits:   ', len(my_l)
        total = sum(int(arg) for arg in my_l)
        print 'sum of digits: ', total
        print '\n'

#end
