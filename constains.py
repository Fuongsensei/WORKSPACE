#pylint:disable = all
from blessed import Terminal
import threading as th 

macro_done : th.Event = th.Event()
get_user_and_path : th.Event = th.Event()
is_run_macro : th.Event = th.Event()
done : th.Event = th.Event()
is_event : th.Event = th.Event()
progress : int = 0
term = Terminal()
