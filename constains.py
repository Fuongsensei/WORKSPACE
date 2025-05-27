#pylint:disable = all
from blessed import Terminal
import threading as th 

done : th.Event = th.Event()
is_event : th.Event = th.Event()
progress : int = 0
term = Terminal()
