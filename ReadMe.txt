StringBuilderBench what is it?

A:
This is a benchmark program which tests the speed of different 
implementaionts of the class StringBuilder in VBC
you can find many at least two different are: 
* Jost Schwider's Concat
* Steve McMahon's StringBuilder

Though the one using MidB is definitely faster than the one using RtlMoveMemory.  

But Jost Schwiders algorithm has one obvious catch in it.
So the speed could be increased about the factor 7 just by 
using a faster string allocation method.

                                 