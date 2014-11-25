#
# This file is part of Dragonfly.
# (c) Copyright 2007, 2008 by Christo Butcher
# Licensed under the LGPL.
#
#   Dragonfly is free software: you can redistribute it and/or modify it 
#   under the terms of the GNU Lesser General Public License as published 
#   by the Free Software Foundation, either version 3 of the License, or 
#   (at your option) any later version.
#
#   Dragonfly is distributed in the hope that it will be useful, but 
#   WITHOUT ANY WARRANTY; without even the implied warranty of 
#   MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU 
#   Lesser General Public License for more details.
#
#   You should have received a copy of the GNU Lesser General Public 
#   License along with Dragonfly.  If not, see 
#   <http://www.gnu.org/licenses/>.
#

"""
This file implements a Win32 keyboard interface using sendinput.

"""


import time
import win32con
from ctypes import windll, c_char, c_wchar
from dragonfly.actions.sendinput import (KeyboardInput, make_input_array,
                                         send_input_array)


#---------------------------------------------------------------------------
# Typeable class.

class Typeable(object):

    __slots__ = ("_code", "_modifiers", "_name")

    def __init__(self, code, modifiers=(), name=None):
        self._code = code
        self._modifiers = modifiers
        self._name = name

    def __str__(self):
        return "%s(%s)" % (self.__class__.__name__, self._name) + repr(self.events())

    def on_events(self, timeout=0):
        events = [(m, True, 0) for m in self._modifiers]
        events.append((self._code, True, timeout))
        return events

    def off_events(self, timeout=0):
        events = [(m, False, 0) for m in self._modifiers]
        events.append((self._code, False, timeout))
        events.reverse()
        return events

    def events(self, timeout=0):
        events = [(self._code, True, 0), (self._code, False, timeout)]
        for m in self._modifiers[-1::-1]:
            events.insert(0, (m, True, 0))
            events.append((m, False , 0))
        return events


#---------------------------------------------------------------------------
# Keyboard access class.

class Keyboard(object):

    shift_code =    win32con.VK_SHIFT
    ctrl_code =     win32con.VK_CONTROL
    alt_code =      win32con.VK_LMENU
    #alt_code =      164
    
    # Neo2 Modifier
    mod2_code = win32con.VK_SHIFT
    mod3_code = 226
    # mod3_code = win32con.VK_CAPITAL
    mod4_code = 223
    # mod4_code = win32con.VK_OEM_8

    @classmethod
    def send_keyboard_events(cls, events):
        """
            Send a sequence of keyboard events.

            Positional arguments:
            events -- a sequence of 3-tuples of the form
                (keycode, down, timeout), where
                keycode (int): virtual key code.
                down (boolean): True means the key will be pressed down,
                    False means the key will be released.
                timeout (int): number of seconds to sleep after
                    the keyboard event.

        """
        items = []
        for keycode, down, timeout in events:
            input = KeyboardInput(keycode, down)
            items.append(input)
            if timeout:
                array = make_input_array(items)
                items = []
                send_input_array(array)
                time.sleep(timeout)
        if items:
            array = make_input_array(items)
            send_input_array(array)
            if timeout: time.sleep(timeout)

    @classmethod
    def xget_virtual_keycode(cls, char):
        print "Warning! Using function that is not covered by NEO2 modifications!"
        
        if isinstance(char, str):
            code = windll.user32.VkKeyScanA(c_char(char))
        else:
            code = windll.user32.VkKeyScanW(c_wchar(char))
        if code == -1:
            raise ValueError("Unknown char: %r" % char)

        # Construct a list of the virtual key code and modifiers.
        codes = [code & 0x00ff]
        if   code & 0x0100: codes.append(cls.shift_code)
        elif code & 0x0200: codes.append(cls.ctrl_code)
        elif code & 0x0400: codes.append(cls.alt_code)
        return codes

    @classmethod
    def get_keycode_and_modifiers(cls, char):
        if isinstance(char, str):
            code = windll.user32.VkKeyScanA(c_char(char))
        else:
            code = windll.user32.VkKeyScanW(c_wchar(char))
        if code == -1:
            raise ValueError("Unknown char: %r" % char)

        # Construct a list of the virtual key code and modifiers.
        modifiers = []
        if   code & 0x0100: modifiers.append(cls.shift_code)
        elif code & 0x0200: modifiers.append(cls.ctrl_code)
        elif code & 0x0400: modifiers.append(cls.alt_code)
        code &= 0x00ff
        return code, modifiers

    # Static Neo2 Typeables
    @classmethod
    def get_neo2_typeable(cls, char):
        return {
            ' ': (32, []),
            'a': (65, []),
            'b': (66, []),
            'c': (67, []),
            'd': (68, []),
            'e': (69, []),
            'f': (70, []),
            'g': (71, []),
            'h': (72, []),
            'i': (73, []),
            'j': (74, []),
            'k': (75, []),
            'l': (76, []),
            'm': (77, []),
            'n': (78, []),
            'o': (79, []),
            'p': (80, []),
            'q': (81, []),
            'r': (82, []),
            's': (83, []),
            't': (84, []),
            'u': (85, []),
            'v': (86, []),
            'w': (87, []),
            'x': (88, []),
            'y': (89, []),
            'z': (90, []),
            'A': (321, [cls.mod2_code]),
            'B': (322, [cls.mod2_code]),
            'C': (323, [cls.mod2_code]),
            'D': (324, [cls.mod2_code]),
            'E': (325, [cls.mod2_code]),
            'F': (326, [cls.mod2_code]),
            'G': (327, [cls.mod2_code]),
            'H': (328, [cls.mod2_code]),
            'I': (329, [cls.mod2_code]),
            'J': (330, [cls.mod2_code]),
            'K': (331, [cls.mod2_code]),
            'L': (332, [cls.mod2_code]),
            'M': (333, [cls.mod2_code]),
            'N': (334, [cls.mod2_code]),
            'O': (335, [cls.mod2_code]),
            'P': (336, [cls.mod2_code]),
            'Q': (337, [cls.mod2_code]),
            'R': (338, [cls.mod2_code]),
            'S': (339, [cls.mod2_code]),
            'T': (340, [cls.mod2_code]),
            'U': (341, [cls.mod2_code]),
            'V': (342, [cls.mod2_code]),
            'W': (343, [cls.mod2_code]),
            'X': (344, [cls.mod2_code]),
            'Y': (345, [cls.mod2_code]),
            'Z': (346, [cls.mod2_code]),
            '0': (48, []),
            '1': (49, []),
            '2': (2236, [cls.mod4_code]),
            '3': (51, []),
            '4': (52, []),
            '5': (53, []),
            '6': (2132, [cls.mod4_code]),
            '7': (55, []),
            '8': (56, []),
            '9': (57, []),
            '!': (4171, [cls.mod3_code]),
            '@': (4185, [cls.mod3_code]),
            '#': (4316, [cls.mod3_code]),
            '$': (4317, [cls.mod3_code]),
            '%': (4173, [cls.mod3_code]),
            '^': (186, [cls.mod3_code]),
            '&': (4177, [cls.mod3_code]),
            '*': (2096, [cls.mod4_code]),
            '(': (4174, [cls.mod3_code]),
            ')': (4178, [cls.mod3_code]),
            '-': (189, [cls.mod4_code]),
            '_': (4182, [cls.mod3_code]),
            '+': (2129, [cls.mod4_code]),
            '`': (191, [cls.mod3_code]),
            '~': (4176, [cls.mod3_code]),
            '[': (4172, [cls.mod3_code]),
            ']': (4163, [cls.mod3_code]),
            '{': (4161, [cls.mod3_code]),
            '}': (4165, [cls.mod3_code]),
            '\\': (4181, [cls.mod3_code]),
            '|': (4318, [cls.mod3_code]),
            ':': (4164, [cls.mod3_code]),
            ';': (4170, [cls.mod3_code]),
            '\'': (4286, [cls.mod3_code]),
            '"': (4284, [cls.mod3_code]),
            ',': (2116, [cls.mod4_code]),
            #            '.': (6330, [cls.mod4_code]),
            '.': (190, []),
            '/': (2105, [cls.mod4_code]),
            '<': (4168, [cls.mod3_code]),
            '>': (4167, [cls.mod3_code]),
            '?': (4179, [cls.mod3_code]),
            '=': (4166, [cls.mod3_code])
        }.get(char, (-1, -1))

    @classmethod
    def get_typeable(cls, char):
        # Call cls.get_neo2_typeable(char) to get static NEO2 Typeables
        code, modifiers = cls.get_neo2_typeable(char)

        if code == -1 and modifiers == -1:
            code, modifiers = cls.get_keycode_and_modifiers(char)
        else:
            code &= 0x00ff
        return Typeable(code, modifiers)


keyboard = Keyboard()
