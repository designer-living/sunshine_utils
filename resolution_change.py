import win32api
import win32.lib.win32con as win32con
import pywintypes
from time import sleep

import argparse

PRESETS = {
    '4k': (3840, 2160),
    '2k': (2560, 1440),
    '1080p': (1920, 1080),
    '800p': (1280, 800),
    '720p': (1280, 720),
}


class ResolutionChanger:

    def __init__(self):
        self.displays = {}
        self._update_display_devices()

    def _update_display_devices(self):
        """
        Enumerates the displays on this machine and updates the local variable with the information.
        :return:
        """
        self.displays = {}
        i = 0
        while True:
            try:
                d = win32api.EnumDisplayDevices(None, i)
                enabled = d.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP == win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP
                if enabled:
                    self.displays[d.DeviceName] = {
                        'device': d,
                        'settings': win32api.EnumDisplaySettings(d.DeviceName, win32con.ENUM_CURRENT_SETTINGS)
                    }
            except pywintypes.error as e:
                break
            finally:
                i += 1

    def _get_primary_display_name(self):
        """
        Loops through the connected displays and finds the primary display.
        Requires self._update_display_devices() to have been called first.
        :return:
        """
        for display_name in self.displays.keys():
            if self.is_primary(display_name):
                return display_name
        return None

    def is_enabled(self, display_name):
        details_map = self.displays[display_name]
        device = details_map['device']
        enabled = device.StateFlags & win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP == win32con.DISPLAY_DEVICE_ATTACHED_TO_DESKTOP
        return enabled

    def is_primary(self, display_name):
        details_map = self.displays[display_name]
        device = details_map['device']
        primary = device.StateFlags & win32con.DISPLAY_DEVICE_PRIMARY_DEVICE == win32con.DISPLAY_DEVICE_PRIMARY_DEVICE
        return primary

    def print_display_details(self):
        """
        Loops through the connected displays and prints info about each display
        Requires self._update_display_devices() to have been called first.
        :return:
        """
        for (display_name, details_map) in self.displays.items():
            settings = details_map['settings']
            enabled = self.is_enabled(display_name)
            primary = self.is_primary(display_name)
            width = "NOT SET"
            height = "NOT SET"
            if settings:
                width = settings.PelsWidth
                height = settings.PelsHeight
            print(f"Display name: {display_name} -  Enabled: {enabled}, Primary: {primary} - {width}x{height}")

    def change_display(self, target_display, new_width, new_height, new_refresh_rate=None, detach_others=False):
        """
        Changes the given display to the given width/height/refresh rate and sets it to primary.
        Requires self._update_display_devices() to have been called first.
        """
        set_primary = False

        if target_display == "primary":
            target_display = self._get_primary_display_name()

        pos = (
            self.displays[target_display]['settings'].Position_x,
            self.displays[target_display]['settings'].Position_y)

        if new_refresh_rate is None:
            print("Switching resolution to {0}x{1}".format(new_width, new_height))
        else:
            print("Switching resolution to {0}x{1} at {2}Hz".format(new_width, new_height, new_refresh_rate))

        devmode = win32api.EnumDisplaySettings(target_display, win32con.ENUM_CURRENT_SETTINGS)
        devmode.PelsWidth = int(new_width)
        devmode.PelsHeight = int(new_height)
        devmode.Position_x = int(pos[0])
        devmode.Position_y = int(pos[1])
        devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT | win32con.DM_POSITION

        if new_refresh_rate is not None:
            devmode.DisplayFrequency = int(new_refresh_rate)
            devmode.Fields |= win32con.DM_DISPLAYFREQUENCY

        win32api.ChangeDisplaySettingsEx(
            target_display, devmode, win32con.CDS_SET_PRIMARY if set_primary else 0)

        if detach_others:
            for display_name in self.displays.keys():
                if display_name != target_display:
                    self.detach_display(display_name)

    def detach_display(self, target_display):
        """
        Changes the given display to the given width/height/refresh rate and sets it to primary.
        Requires self._update_display_devices() to have been called first.
        """
        if target_display == "primary":
            target_display = self._get_primary_display_name()

        devmode = win32api.EnumDisplaySettings(target_display, win32con.ENUM_CURRENT_SETTINGS)
        devmode.PelsWidth = 0
        devmode.PelsHeight = 0
        devmode.Fields = win32con.DM_PELSWIDTH | win32con.DM_PELSHEIGHT | win32con.DM_POSITION

        win32api.ChangeDisplaySettingsEx(target_display, devmode)

    def reset_resolutions(self):
        """
        Resets the resolution of all monitors
        :return:
        """
        print("Resetting resolution")
        win32api.ChangeDisplaySettings(None, 0)
        print("Reset to:")
        self._update_display_devices()
        self.print_display_details()


if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Changes monitor resolution")

    group = parser.add_mutually_exclusive_group(required=True)

    group.add_argument('-l', '--list-displays', default=False, action="store_true",
                       help='List current displays and exit')

    group.add_argument('--width', type=int, help='Resolution Width', required=False)
    parser.add_argument('--height', type=int, help='Resolution Height', required=False)

    parser.add_argument('--refreshrate', type=int, required=False,
                        help=f'The refresh rate in hertz')

    parser.add_argument('-d', '--display', default='primary', type=str, required=False,
                        help=r'Which display to use e.g. \\.\DISPLAY1 (defaults to current primary)')

    parser.add_argument('--detach-other-monitors', default=False, action="store_true",
                       help="Detach other monitors except the named one")

    group.add_argument('-p', '--preset', type=str, required=False,
                       help=f'A preset to use can be one of: {PRESETS.keys()}')

    group.add_argument('-r', '--reset', default=False, action="store_true",
                       help='Reset the resolution that a previous run has changed')

    group.add_argument('--detach-monitor', default=False, action="store_true",
                       help="Detach the named display")

    parser.add_argument('--wait', type=int, default=5,
                        help="Time in seconds to wait after changing resolution before exiting")

    parser.add_argument('--debug', default=False, action="store_true",
                        help="Enable debug mode so the program won't exit after running")

    args = parser.parse_args()
    resolution_changer = ResolutionChanger()

    if args.list_displays:
        resolution_changer.print_display_details()

    elif args.width and not args.height:
        print("If using width you must also specfiy a height")
        parser.print_help()
        exit(-1)
    else:
        if args.width:
            resolution_changer.change_display(args.display, args.width, args.height, args.refreshrate,
                                              args.detach_other_monitors)
        elif args.preset:
            (width, height) = PRESETS.get(args.preset)
            resolution_changer.change_display(args.display, width, height, args.refreshrate,
                                              args.detach_other_monitors)
        elif args.reset:
            resolution_changer.reset_resolutions()
        elif args.detach_monitor:
            resolution_changer.detach_display(args.display)
        else:
            print("Unexpected options")
            parser.print_help()

    sleep(args.wait)

    if args.debug:
        # Leave window open for debugging
        input("Paused for debug review. Press Enter key to close.")
