import autopy
import math
import time
import random

from threading import Timer


TWO_PI = math.pi * 2.0

rand_minutes = random.randrange(60, 90) * 240


timeout = time.time() + (480 * 60)


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
        autopy.mouse.move(x, y)
        time.sleep(random.uniform(0.001, 0.003))
        print(time.time())


try:
    while time.time() < timeout:
        sine_mouse_wave()
except KeyboardInterrupt:
    pass
