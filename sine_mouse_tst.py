import autopy
import math
import time
import random

TWO_PI = math.pi * 2.0

rand_minutes = random.randrange(45, 55) * 600


timeout = time.time() + rand_minutes + random.randrange(5, 60)


def sine_mouse_wave():
    """
    Moves the mouse in a sine wave from the left edge of
    the screen to the right.
    """    
    autopy.key.tap(autopy.key.Code.SHIFT, [autopy.key.Modifier.META])

    width, height = autopy.screen.size()
    height /= 2
    height -= 10  # Stay in the screen bounds.

    for x in range(int(width)):
        # y = int(height * math.sin((TWO_PI * x) / width) + height)
        y = int(height * math.sin((TWO_PI * x) / width) + height)
        print(y)
        autopy.mouse.move(x, y)
        time.sleep(random.uniform(0.001, 0.003))


try:
    while time.time() < timeout:
        autopy.key.tap(autopy.key.Code.SHIFT, [autopy.key.Modifier.META])        
        sine_mouse_wave()
except KeyboardInterrupt:
    pass
