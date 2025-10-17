import cmd
import sys
import os
import csv
import time
import readline
import pyvisa
import pyvisa_py

# Optional: history file support
HISTORY_FILE = os.path.expanduser("~/.instrument_tool_history")

class Console(cmd.Cmd):
    """
    A command-line interface console for controlling IVI VISA instruments.

    This class handles instrument selection, connection, and provides a set of
    high-level commands for common instrument operations.

    :ivar rm: The PyVISA ResourceManager instance.
    :vartype rm: pyvisa.ResourceManager or None
    :ivar selected_device_id: The VISA resource string of the currently connected instrument.
    :vartype selected_device_id: str or None
    :ivar instrument: The PyVISA resource object for the connected instrument.
    :vartype instrument: pyvisa.Resource or None
    """
    intro = '''
    Welcome! This is a command prompt for controlling IVI VISA instruments.
    Type 'help' or '?' to list commands. Type 'exit' to quit.
    '''
    prompt = ">>> "

    # Resource manager and selected instrument
    try:
        rm = pyvisa.ResourceManager()
    except Exception as e:
        # If ResourceManager fails at import time, set to None and handle later
        print(f"Warning: Could not create ResourceManager at import: {e}")
        rm = None

    selected_device_id = None
    instrument = None

    def __init__(self):
        super().__init__()
        # Load history (if supported)
        try:
            readline.read_history_file(HISTORY_FILE)
        except Exception:
            pass

    # -----------------------
    # Device management
    # -----------------------
    def do_devicelist(self, arg):
        """
        Lists detected devices.

        Retrieves and prints a list of all available VISA resources detected
        by the PyVISA Resource Manager.

        Usage: devicelist
        """
        if not self.rm:
            print("ResourceManager not available.")
            return
        try:
            devices = self.rm.list_resources()
            if devices:
                print("Available VISA Resources:")
                for device in devices:
                    print(f"  {device}")
            else:
                print("No VISA resources detected.")
        except Exception as e:
            print(f"Error listing devices: {e}")

    def do_deviceselect(self, device_id):
        """
        Selects and connects to a particular device.

        Closes any currently open connection, attempts to open the specified
        device ID, sets a default timeout (5000 ms), and queries the IDN.

        :param device_id: The VISA resource string of the device to connect to (e.g., "GPIB0::2::INSTR").
        :type device_id: str
        :raises pyvisa.errors.VisaIOError: If connection to the device fails.
        :raises Exception: For other unexpected errors.

        Usage: deviceselect \"device_id\"
        Example: deviceselect \"GPIB0::2::INSTR\"
        """
        if not device_id:
            print("Please provide a device ID. Use 'devicelist' to see options.")
            return
        if not self.rm:
            print("ResourceManager not available.")
            return

        device_id_clean = device_id.strip().strip('"')
        try:
            # Close previous instrument if open
            if self.instrument:
                try:
                    self.instrument.close()
                except Exception:
                    pass
                self.instrument = None
                self.selected_device_id = None

            # Open new device
            self.instrument = self.rm.open_resource(device_id_clean)
            self.selected_device_id = device_id_clean
            # Configure some sensible defaults (timeouts)
            try:
                # 5 second timeout by default
                self.instrument.timeout = getattr(self.instrument, 'timeout', 5000)
            except Exception:
                pass

            print(f"Successfully selected and connected to: {self.selected_device_id}")
            # Optionally show ID
            try:
                self.do_id(None)
            except Exception:
                pass
        except pyvisa.errors.VisaIOError as e:
            print(f"Error connecting to device '{device_id_clean}': {e}")
            self.instrument = None
            self.selected_device_id = None
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_deviceinfo(self, device_id):
        """
        Shows a basic info hint for a device string (without connecting).

        :param device_id: The VISA resource string to display information for.
        :type device_id: str

        Usage: deviceinfo \"device_id\"
        Note: For manufacturer/model use deviceselect then id.
        """
        if not device_id:
            print("Usage: deviceinfo \"device_id\"")
            return
        print(f"Device ID: {device_id.strip().strip('\"')}")
        print("To get manufacturer/model, use 'deviceselect' then 'id' (which queries *IDN?).")

    def do_id(self, arg):
        """
        Queries the selected instrument for its identification string (*IDN?).

        :raises pyvisa.errors.VisaIOError: If the IDN query fails.

        Usage: id
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        try:
            idn = self.instrument.query('*IDN?')
            print(f"Instrument IDN: {idn.strip()}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error querying *IDN?: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    # -----------------------
    # Generic SCPI Commands
    # -----------------------
    def do_write(self, command):
        """
        Sends a raw SCPI command to the selected instrument (no response expected).

        :param command: The SCPI command string to send.
        :type command: str

        :raises pyvisa.errors.VisaIOError: If writing the command fails.

        Usage: write \"SCPI_COMMAND\"
        Example: write \"*RST\"
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        if not command:
            print("Usage: write \"SCPI_CMD\"")
            return
        try:
            cmd_str = command.strip().strip('"')
            self.instrument.write(cmd_str)
            print(f"Sent: {cmd_str}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error writing command: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_query(self, command):
        """
        Sends a raw SCPI command and expects a response.

        :param command: The SCPI query string to send (e.g., "SYST:ERR?").
        :type command: str

        :raises pyvisa.errors.VisaIOError: If the query fails or times out.

        Usage: query \"SCPI_QUERY?\"
        Example: query \"SYST:ERR?\"
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        if not command:
            print("Usage: query \"SCPI_QUERY\"")
            return
        try:
            cmd_str = command.strip().strip('"')
            response = self.instrument.query(cmd_str)
            print(f"Response: {response.strip()}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error querying command: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_reset(self, arg):
        """
        Resets the instrument to its factory default state.

        Uses the common '*RST' command. This does not power cycle the device.

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: reset
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        try:
            self.instrument.write('*RST')
            print("Instrument reset command (*RST) sent.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error resetting instrument: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_wait_opc(self, arg):
        """
        Waits for all pending instrument operations to complete.

        Uses the common '*OPC' command (Operation Complete). This command is essential
        when automating sequences where the next command depends on the completion
        of the previous one.

        :raises pyvisa.errors.VisaIOError: If the command fails or times out.

        Usage: wait_opc
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        try:
            # Use query with *OPC? to force synchronous operation in PyVISA
            self.instrument.query('*OPC?')
            print("Operation Complete (*OPC?) verified.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error waiting for OPC: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_get_error(self, arg):
        """
        Queries the instrument's Standard Event Status Register and Error Queue.

        Uses the common ':SYSTem:ERRor?' command.

        :raises pyvisa.errors.VisaIOError: If the query fails.

        Usage: get_error
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        try:
            error = self.instrument.query(':SYSTem:ERRor?')
            print(f"Instrument Error: {error.strip()}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error querying system error: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    # -----------------------
    # DMM Specific Commands
    # -----------------------
    def do_dmm_func_set(self, func):
        """
        Sets the primary measurement function on the DMM.

        Uses the common ':SENSe:FUNCtion' command.

        :param func: The DMM function string (e.g., "VOLT:DC", "CURR:AC", "RES").
        :type func: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: dmm_func_set \"function\"
        Example: dmm_func_set \"VOLT:DC\"
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        if not func:
            print("Usage: dmm_func_set \"FUNCTION\"")
            return
        try:
            scpi_func = func.strip().strip('"').upper()
            # Some instruments expect different syntax; using a common form
            self.instrument.write(f':SENSe:FUNCtion "{scpi_func}"')
            print(f"DMM function set to: {scpi_func}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting DMM function: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_range_set(self, range_val):
        """
        Sets the measurement range for the currently selected function.

        Uses the common ':SENSe:RANGe' command.

        :param range_val: The desired measurement range value (float) or "AUTO".
        :type range_val: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: dmm_range_set <range_value | AUTO>
        Example: dmm_range_set 10
                 dmm_range_set AUTO
        """
        if not self.instrument:
            print("No device selected. Use 'deviceselect' first.")
            return
        if not range_val:
            print("Usage: dmm_range_set <range_value|AUTO>")
            return
        try:
            self.instrument.write(f':SENSe:RANGe {range_val.strip()}')
            print(f"DMM range set to: {range_val.strip()}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting DMM range: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_dmm_autoranging(self, state):
        """
        Sets the DMM to use manual or auto-ranging.

        Uses the common ':SENSe:RANGe:AUTO' command. This usually applies to the
        currently selected function (VOLT:DC, CURR:AC, etc.).

        :param state: The auto-ranging state ("ON" or "OFF").
        :type state: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: dmm_autoranging <ON|OFF>
        Example: dmm_autoranging ON
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not state:
            print("Usage: dmm_autoranging <ON|OFF>")
            return
        try:
            scpi_state = state.upper()
            if scpi_state not in ["ON", "OFF", "1", "0"]:
                print("Invalid state. Use ON or OFF.")
                return
            self.instrument.write(f':SENSe:RANGe:AUTO {scpi_state}')
            print(f"DMM auto-ranging set to: {scpi_state}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting DMM auto-ranging: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_delay_set(self, delay):
        """
        Sets the measurement integration time or delay.

        Uses the common ':SENSe:DELay' command (often in seconds).

        :param delay: The desired measurement delay (float, in seconds).
        :type delay: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: dmm_delay_set <seconds>
        Example: dmm_delay_set 0.1
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not delay:
            print("Usage: dmm_delay_set <seconds>")
            return
        try:
            self.instrument.write(f':SENSe:DELay {delay.strip()}')
            print(f"DMM measurement delay set to: {delay.strip()} s")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting DMM delay: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_resolution_set(self, resolution):
        """
        Sets the measurement resolution.

        Uses the common ':SENSe:RESolution' command.

        :param resolution: The desired resolution value (float).
        :type resolution: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: dmm_resolution_set <resolution_value>
        Example: dmm_resolution_set 0.001
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not resolution:
            print("Usage: dmm_resolution_set <resolution>")
            return
        try:
            self.instrument.write(f':SENSe:RESolution {resolution.strip()}')
            print(f"DMM resolution set to: {resolution.strip()}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting DMM resolution: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    # --- DMM Measurement Commands ---
    def do_dmm_measure_dc_v(self, arg):
        """
        Measures DC Voltage (V).

        Uses the common ':MEASure:VOLTage:DC?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure_dc_v
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            measurement = self.instrument.query(':MEASure:VOLTage:DC?')
            print(f"DC Voltage: {measurement.strip()} V")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring DC Voltage. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_measure_ac_v(self, arg):
        """
        Measures AC Voltage (RMS).

        Uses the common ':MEASure:VOLTage:AC?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure_ac_v
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            measurement = self.instrument.query(':MEASure:VOLTage:AC?')
            print(f"AC Voltage (RMS): {measurement.strip()} V")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring AC Voltage. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_measure_dc_i(self, arg):
        """
        Measures DC Current.

        Uses the common ':MEASure:CURRent:DC?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure_dc_i
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            measurement = self.instrument.query(':MEASure:CURRent:DC?')
            print(f"DC Current: {measurement.strip()} A")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring DC Current. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_measure_ac_i(self, arg):
        """
        Measures AC Current (RMS).

        Uses the common ':MEASure:CURRent:AC?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure_ac_i
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            measurement = self.instrument.query(':MEASure:CURRent:AC?')
            print(f"AC Current (RMS): {measurement.strip()} A")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring AC Current. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_dmm_measure_continuity(self, arg):
        """
        Measures continuity and returns the resistance reading.

        This test often involves selecting the 'CONTinuity' function and then reading
        the resistance. Most DMMs report resistance during this test.

        Uses the common ':MEASure:CONTinuity?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure_continuity
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            # Some DMMs use this dedicated command, others require setting the function first.
            measurement = self.instrument.query(':MEASure:CONTinuity?')
            # Note: The raw measurement is typically resistance, in Ohms.
            print(f"Continuity Resistance: {measurement.strip()} Ohm")
            # For automation, users would check if the resistance is below a threshold (e.g., < 10 Ohm).
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring continuity. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
    
    def do_dmm_measure_diode(self, arg):
        """
        Measures the forward voltage drop across a diode.

        This test typically involves selecting the 'DIODe' function and then reading
        the voltage drop (in Volts).

        Uses the common ':MEASure:DIODe?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure_diode
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            measurement = self.instrument.query(':MEASure:DIODe?')
            # The measurement is the forward voltage drop, typically in Volts.
            print(f"Diode Forward Voltage: {measurement.strip()} V")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring diode forward voltage. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_measure2_resistance(self, arg):
        """
        Measures Resistance (2-wire).

        Uses the common ':MEASure:RESistance?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure2_resistance
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            measurement = self.instrument.query(':MEASure:RESistance?')
            print(f"Resistance: {measurement.strip()} Ohm")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring 2-wire Resistance. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_dmm_measure4_resistance(self, arg):
        """
        Measures Resistance (4-wire, Kelvin).

        Uses the common ':MEASure:FRESistance?' command.

        :raises pyvisa.errors.VisaIOError: If measurement fails.

        Usage: dmm_measure4_resistance
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            # Perform a 4-wire resistance measurement
            measurement = self.instrument.query(':MEASure:FRESistance?')
            print(f"4-Wire Resistance: {measurement.strip()} Ohm")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error measuring 4-wire Resistance. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")


    # -----------------------
    # Oscilloscope Commands
    # -----------------------
    def do_oscope_set_timebase(self, sec_per_div):
        """
        Sets the horizontal scale (Time per Division).

        Uses the common ':HORizontal:SCAle' command.

        :param sec_per_div: Time per division in seconds (float).
        :type sec_per_div: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: oscope_set_timebase <seconds_per_div>
        Example: oscope_set_timebase 0.005
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not sec_per_div:
            print("Usage: oscope_set_timebase <seconds_per_div>")
            return
        try:
            self.instrument.write(f':HORizontal:SCAle {sec_per_div}')
            print(f"Oscilloscope timebase set to: {sec_per_div} s/div")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting timebase: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_set_vertscale(self, channel_volt_per_div):
        """
        Sets the vertical scale (Volts per Division) for a specified channel.

        Uses the common ':CHANnel<n>:SCAle' command.

        :param channel_volt_per_div: Comma-separated string of channel number and volts per division (e.g., "1,0.5").
        :type channel_volt_per_div: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: oscope_set_vertscale <channel>,<volts_per_div>
        Example: oscope_set_vertscale 1,0.5
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not channel_volt_per_div:
            print("Usage: oscope_set_vertscale <channel>,<volts_per_div>")
            return
        try:
            channel, volts_per_div = map(str.strip, channel_volt_per_div.split(','))
            self.instrument.write(f':CHANnel{channel}:SCAle {volts_per_div}')
            print(f"Oscilloscope Channel {channel} scale set to: {volts_per_div} V/div")
        except ValueError:
            print("Invalid format. Use: <channel>,<volts_per_div> (e.g., 1,0.5)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting vertical scale: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_measure_param(self, channel_parameter):
        """
        Reads a measured parameter from the oscilloscope.

        Uses the common ':MEASure:<PARAM>? CHANnel<n>' command.

        :param channel_parameter: Comma-separated string of channel number and parameter (e.g., "1,FREQ").
        :type channel_parameter: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If measurement is unsupported or fails.

        Usage: oscope_measure_param <channel>,<parameter>
        Example: oscope_measure_param 1,FREQ
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not channel_parameter:
            print("Usage: oscope_measure_param <channel>,<parameter>")
            return
        try:
            channel, parameter = map(str.strip, channel_parameter.split(','))
            parameter_scpi = parameter.upper()
            raw_query = f':MEASure:{parameter_scpi}? CHANnel{channel}'
            measurement = self.instrument.query(raw_query)
            print(f"Channel {channel} {parameter_scpi}: {measurement.strip()}")
        except ValueError:
            print("Invalid format. Use: <channel>,<parameter> (e.g., 1,FREQ)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error reading measurement: {e}. Check if this measurement is supported.")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_set_trigger_source(self, source):
        """
        Sets the signal source for the main trigger.

        Uses the common ':TRIGger:SOURce' command.

        :param source: The trigger source (e.g., "CHAN1", "EXT").
        :type source: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: oscope_set_trigger_source <channel_or_ext>
        Example: oscope_set_trigger_source CHAN1
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not source:
            print("Usage: oscope_set_trigger_source <source>")
            return
        try:
            self.instrument.write(f':TRIGger:SOURce {source.upper()}')
            print(f"Oscilloscope trigger source set to: {source.upper()}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting trigger source: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_set_trigger_level(self, channel_level):
        """
        Sets the trigger level voltage for a specific channel's trigger settings.

        This command is typically applied to the currently selected trigger source
        (set via :TRIGger:SOURce). Many instruments only require the level, but
        this format allows for explicit channel targeting if required.

        Uses the common ':TRIGger:LEVel' command (applied globally or to current source).

        :param channel_level: Comma-separated string of **channel number** and the trigger level voltage (e.g., "1,0.5").
        :type channel_level: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: oscope_set_trigger_level <channel>,<voltage>
        Example: oscope_set_trigger_level 1,0.5
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not channel_level:
            print("Usage: oscope_set_trigger_level <channel>,<voltage>")
            return
        try:
            channel, level = map(str.strip, channel_level.split(','))
            self.instrument.write(f':TRIGger:LEVel {level}')
            print(f"Oscilloscope trigger level set to: {level} V (Assumed for Channel {channel} source)")
        except ValueError:
            print("Invalid format. Use: <channel>,<voltage> (e.g., 1,0.5)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting trigger level: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_set_trigger_slope(self, channel_slope):
        """
        Sets the trigger slope (edge) for a specific channel (source).

        Uses the common ':TRIGger:EDGE:SLOpe' command (applied globally or to current source).

        :param channel_slope: Comma-separated string of **channel number** and the trigger slope ("POS", "NEG", or "EITH").
        :type channel_slope: str

        :raises ValueError: If the input format is incorrect or slope is invalid.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: oscope_set_trigger_slope <channel>,<POS|NEG|EITH>
        Example: oscope_set_trigger_slope 1,POS
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not channel_slope:
            print("Usage: oscope_set_trigger_slope <channel>,<POS|NEG|EITH>")
            return
        try:
            channel, slope = map(str.strip, channel_slope.split(','))
            scpi_slope = slope.upper()
            if scpi_slope not in ["POS", "NEG", "EITH"]:
                print("Invalid slope. Use POS, NEG, or EITH.")
                return
            # Assuming the scope uses the global trigger system or that setting the slope
            # applies to the active source/channel.
            self.instrument.write(f':TRIGger:EDGE:SLOpe {scpi_slope}')
            print(f"Oscilloscope trigger slope set to: {scpi_slope} (Assumed for Channel {channel} source)")
        except ValueError:
            print("Invalid format. Use: <channel>,<POS|NEG|EITH> (e.g., 1,POS)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting trigger slope: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_get_setup(self, arg):
        """
        Gets the current instrument setup string (useful for saving/restoring setup).

        Uses the common ':SETup?' query.

        :raises pyvisa.errors.VisaIOError: If the query fails.

        Usage: oscope_get_setup
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            setup = self.instrument.query(':SETup?')
            print(f"Oscilloscope Setup String (truncated to 200 chars): {setup.strip()[:200]}...")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error getting setup. Check instrument capability: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_run(self, arg):
        """
        Sets the oscilloscope to run/continuous mode.

        Uses the common ':RUN' command.

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: oscope_run
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            self.instrument.write(':RUN')
            print("Oscilloscope set to RUN mode.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting RUN mode: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_oscope_screen_capture(self, filename_format):
        """
        Captures the oscilloscope's screen and saves the image data to a local file (PNG or JPEG).

        Uses common screen dump commands like ':DISPlay:DATA?' or ':HARDCopy:DATA?'.

        :param filename_format: The local file path and requested format (e.g., "scope_image.png").
        :type filename_format: str

        :raises pyvisa.errors.VisaIOError: If the instrument communication fails.

        Usage: oscope_screen_capture <filename.png>
        Example: oscope_screen_capture transient_fault.jpeg
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not filename_format:
            print("Usage: oscope_screen_capture <filename.png|filename.jpeg>")
            return
        
        try:
            filename = filename_format.strip()
            extension = filename.split('.')[-1].upper()

            if extension in ["PNG", "JPEG", "JPG"]:
                scpi_format = extension
                # Common Oscilloscope SCPI for image transfer
                SCPI_COMMAND = f':DISPlay:DATA? {scpi_format}' 
                
                print(f"Requesting scope screen data ({scpi_format})...")

                image_data = self.instrument.query_binary_values(
                    SCPI_COMMAND, 
                    datatype='s',
                    container=bytes,
                    is_termination_char=False
                )
                
                if not image_data:
                    print("Error: Received no image data from the instrument.")
                    return

                with open(filename, 'wb') as f:
                    f.write(image_data)
                
                print(f"Scope screen capture successfully saved to: {filename}")
            else:
                print("Unsupported format for screen capture. Use PNG or JPEG/JPG.")

        except pyvisa.errors.VisaIOError as e:
            print(f"Error communicating with instrument: {e}")
            print("HINT: Ensure the waveform source is set via 'oscope_set_wfm_source'.")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_oscope_capture_data(self, filename_format):
        """
        Captures and processes the raw waveform data from the selected channel and saves it as TXT or CSV.

        This process requires querying the waveform preamble (scale factors) and the data points,
        then converting the raw digital values into real-world Volts and Seconds.

        :param filename_format: The local file path and requested format (e.g., "data.csv").
        :type filename_format: str

        :raises pyvisa.errors.VisaIOError: If communication fails.

        Usage: oscope_capture_data <filename.csv>
        Example: oscope_capture_data sine_wave.txt
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not filename_format:
            print("Usage: oscope_capture_data <filename.csv|filename.txt>")
            return

        try:
            filename = filename_format.strip()
            extension = filename.split('.')[-1].upper()

            if extension not in ["CSV", "TXT"]:
                print("Unsupported format for data capture. Use CSV or TXT.")
                return

            # --- 1. Get Preamble (Scaling Factors) ---
            # Assume user has set waveform source via do_oscope_set_wfm_source
            preamble_str = self.instrument.query(':WAVeform:PREamble?')
            preamble = list(map(float, preamble_str.strip().split(',')))
            
            # Preamble indices for common SCPI devices:
            # 5: X_INCrement (Time/Point)
            # 7: Y_INCrement (Volts/Digit)
            # 8: Y_Reference (Vertical Offset)
            X_INCrement = preamble[5]
            Y_INCrement = preamble[7]
            Y_Reference = preamble[8]

            # --- 2. Get Raw Data ---
            self.instrument.write(':WAVeform:FORMat ASCii') # Set data format to ASCII
            raw_data_str = self.instrument.query(':WAVeform:DATA?')
            raw_data = list(map(float, raw_data_str.strip().split(',')))
            
            # --- 3. Process Data ---
            processed_data = []
            for i, raw_value in enumerate(raw_data):
                # Calculate real voltage: Voltage = Y_INCrement * (Raw_Value - Y_Reference)
                voltage = Y_INCrement * (raw_value - Y_Reference)
                # Calculate time: Time = i * X_INCrement
                time_sec = i * X_INCrement
                processed_data.append((time_sec, voltage))

            # --- 4. Write to File ---
            if extension == "CSV":
                with open(filename, 'w', newline='') as f:
                    writer = csv.writer(f)
                    writer.writerow(['Time (s)', 'Voltage (V)'])
                    writer.writerows(processed_data)
            elif extension == "TXT":
                with open(filename, 'w') as f:
                    f.write("Time (s)\tVoltage (V)\n")
                    for t, v in processed_data:
                        f.write(f"{t:.6e}\t{v:.6e}\n")

            print(f"Waveform data successfully processed and saved to: {filename} ({len(raw_data)} points).")

        except pyvisa.errors.VisaIOError as e:
            print(f"Error communicating with instrument: {e}")
            print("HINT: Ensure the instrument is ready and the waveform source is set.")
        except Exception as e:
            print(f"An unexpected error occurred during data processing: {e}")

    # -----------------------
    # Function Generator (AFG) Commands
    # -----------------------
    def do_afg_set_wave(self, args):
        """
        Sets the function generator waveform, frequency, and amplitude.

        Uses the common ':FUNCtion', ':FREQuency', and ':VOLTage' commands.

        :param args: Comma-separated string of waveform (WAVE), frequency (FREQ), and amplitude (AMPL).
        :type args: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects any command.

        Usage: afg_set_wave \"waveform,frequency,amplitude\"
        Example: afg_set_wave \"SIN,1000,1.0\"
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not args:
            print("Usage: afg_set_wave \"WAVE, FREQ, AMPL\"")
            return
        try:
            waveform, freq, ampl = map(str.strip, args.split(','))
            # Common AFG SCPI commands
            self.instrument.write(f':FUNCtion {waveform.upper()}')
            self.instrument.write(f':FREQuency {freq}')
            self.instrument.write(f':VOLTage {ampl}')
            print(f"AFG set to {waveform.upper()} wave, {freq} Hz, {ampl} Vpp")
        except ValueError:
            print("Invalid format. Use: afg_set_wave \"WAVE,FREQ,AMPL\" (e.g., SIN,1000,1.0)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting waveform: {e}")
        except Exception as e:
            print(f"Unexpected error: {e}")

    def do_afg_output_on(self, arg):
        """
        Turns ON the AFG output.

        Uses the common ':OUTPut ON' command.

        :raises Exception: If command fails.

        Usage: afg_output_on
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            self.instrument.write(':OUTPut ON')
            print("AFG output turned ON.")
        except Exception as e:
            print(f"Error enabling output: {e}")

    def do_afg_output_off(self, arg):
        """
        Turns OFF the AFG output.

        Uses the common ':OUTPut OFF' command.

        :raises Exception: If command fails.

        Usage: afg_output_off
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            self.instrument.write(':OUTPut OFF')
            print("AFG output turned OFF.")
        except Exception as e:
            print(f"Error disabling output: {e}")
            
    def do_afg_psu_slew_set(self, channel_rate):
        """
        Sets the maximum output voltage slew rate (V/s) for a channel.

        This limits how quickly the output voltage can change. Essential for DUT protection.
        This command is highly instrument-specific but commonly uses a format like ':VOLTage:SLEW:RATE'.
        Assumes channel 1 if not explicitly specified.

        :param channel_rate: Comma-separated string of channel (if applicable) and rate (e.g., "1,10.0" or "10.0").
        :type channel_rate: str

        :raises Exception: If command fails or the format is incorrect.

        Usage: afg_psu_slew_set <rate_V_per_s>
        Example: afg_psu_slew_set 10.0
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not channel_rate:
            print("Usage: afg_psu_slew_set <rate_V_per_s>")
            return
        try:
            # Handle potential channel input, defaulting to the full string as the rate if no comma found.
            if ',' in channel_rate:
                channel, rate = map(str.strip, channel_rate.split(','))
                # If supported, use the channel command, e.g., :SOURce1:VOLTage:SLEW:RATE
                self.instrument.write(f':VOLTage:SLEW:RATE {rate}') 
                print(f"Slew rate set to {rate} V/s (Assuming Channel {channel} or single output).")
            else:
                rate = channel_rate.strip()
                self.instrument.write(f':VOLTage:SLEW:RATE {rate}')
                print(f"Slew rate set to {rate} V/s.")

        except Exception as e:
            print(f"Error setting slew rate. Command likely unsupported by instrument: {e}")

    # -----------------------
    # Power Supply (PSU) Commands
    # -----------------------
    def do_psu_set_voltage(self, voltage):
        """
        Sets PSU output voltage.

        Uses the common ':VOLTage' command.

        :param voltage: The desired output voltage (float).
        :type voltage: str

        :raises Exception: If command fails.

        Usage: psu_set_voltage <voltage>
        Example: psu_set_voltage 5.0
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not voltage:
            print("Usage: psu_set_voltage <voltage>")
            return
        try:
            self.instrument.write(f':VOLTage {voltage}')
            print(f"Voltage set to {voltage} V")
        except Exception as e:
            print(f"Error setting voltage: {e}")

    def do_psu_set_current(self, current):
        """
        Sets PSU current limit.

        Uses the common ':CURRent' command.

        :param current: The desired current limit (float).
        :type current: str

        :raises Exception: If command fails.

        Usage: psu_set_current <current>
        Example: psu_set_current 0.5
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not current:
            print("Usage: psu_set_current <current>")
            return
        try:
            self.instrument.write(f':CURRent {current}')
            print(f"Current limit set to {current} A")
        except Exception as e:
            print(f"Error setting current: {e}")

    def do_psu_output_on(self, arg):
        """
        Turns ON the PSU output.

        Uses the common ':OUTPut ON' command.

        :raises Exception: If command fails.

        Usage: psu_output_on
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            self.instrument.write(':OUTPut ON')
            print("Power supply output ON.")
        except Exception as e:
            print(f"Error enabling PSU output: {e}")

    def do_psu_output_off(self, arg):
        """
        Turns OFF the PSU output.

        Uses the common ':OUTPut OFF' command.

        :raises Exception: If command fails.

        Usage: psu_output_off
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            self.instrument.write(':OUTPut OFF')
            print("Power supply output OFF.")
        except Exception as e:
            print(f"Error disabling PSU output: {e}")
            
    def do_psu_set_ovp(self, voltage):
        """
        Sets the Over-Voltage Protection (OVP) limit.

        Once the output voltage exceeds this value, the PSU will shut down (trip).

        Uses the common ':VOLTage:PROTection:LEVel' command.

        :param voltage: The OVP limit voltage (float).
        :type voltage: str

        :raises Exception: If command fails.

        Usage: psu_set_ovp <voltage>
        Example: psu_set_ovp 5.5
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not voltage:
            print("Usage: psu_set_ovp <voltage>")
            return
        try:
            self.instrument.write(f':VOLTage:PROTection:LEVel {voltage}')
            print(f"OVP set to {voltage} V")
        except Exception as e:
            print(f"Error setting OVP: {e}")
            
    def do_psu_set_ocp(self, current):
        """
        Sets the Over-Current Protection (OCP) limit.

        Once the output current exceeds this value, the PSU will shut down (trip).
        Uses the common ':CURRent:PROTection:LEVel' command or similar. Note that
        this differs from the standard current limit command (`:CURRent`), as OCP
        is a hard trip point.

        :param current: The OCP limit current (float).
        :type current: str

        :raises Exception: If command fails.

        Usage: psu_set_ocp <current>
        Example: psu_set_ocp 0.6
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not current:
            print("Usage: psu_set_ocp <current>")
            return
        try:
            # Note: The specific SCPI command for OCP varies. This is a common standard.
            self.instrument.write(f':CURRent:PROTection:LEVel {current}')
            print(f"OCP set to {current} A")
        except Exception as e:
            print(f"Error setting OCP: {e}")
            
    def do_psu_set_otp(self, temperature):
        """
        Sets a software-defined Over-Temperature Protection (OTP) trip point.

        NOTE: OTP is often a fixed hardware limit and non-programmable. This command
        uses a less common SCPI structure, ':SENSe:TEMPerature:PROTection:LEVel',
        and may not be supported by all instruments.

        :param temperature: The OTP limit temperature (float, typically in Celsius).
        :type temperature: str

        :raises Exception: If command fails or is unsupported by the instrument.

        Usage: psu_set_otp <temperature>
        Example: psu_set_otp 85
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not temperature:
            print("Usage: psu_set_otp <temperature>")
            return
        try:
            # This is a speculative SCPI command based on the structure of OVP/OCP.
            self.instrument.write(f':SENSe:TEMPerature:PROTection:LEVel {temperature}')
            print(f"OTP level set to {temperature} °C.")
        except Exception as e:
            print(f"Error setting OTP. Command may be unsupported by instrument: {e}")

    def do_psu_measure_output(self, measure_type):
        """
        Measures the actual output voltage or current.

        Uses the common ':MEASure:<TYPE>?' command.

        :param measure_type: The type of measurement to take ("VOLT", "CURR").
        :type measure_type: str

        :raises Exception: If command fails.

        Usage: psu_measure_output <VOLT|CURR>
        Example: psu_measure_output VOLT
        """
        if not self.instrument:
            print("No device selected.")
            return
        measure_type = measure_type.strip().upper()
        if measure_type not in ["VOLT", "CURR"]:
            print("Usage: psu_measure_output <VOLT|CURR>")
            return
        try:
            measurement = self.instrument.query(f':MEASure:{measure_type}?')
            unit = 'V' if measure_type == 'VOLT' else 'A'
            print(f"Output {measure_type}: {measurement.strip()} {unit}")
        except Exception as e:
            print(f"Error measuring output {measure_type}: {e}")
            
    def do_psu_protection_clear(self, arg):
        """
        Clears the Over-Voltage/Over-Current/Thermal Protection (OVP/OCP/OTP) trip state.

        When a protection mechanism trips, the output is disabled. This command
        is required to clear the trip state and allow the output to be re-enabled
        (via `psu_output_on`).

        Uses the common '*CLS' (Clear Status) or a specific command like ':OUTPut:PROTection:CLEar'.

        :raises Exception: If command fails.

        Usage: psu_protection_clear
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            # Many instruments use a specific command for clearing protection trips,
            # while others might use the generic *CLS. Using the generic clear for robustness.
            # If *CLS is too broad, the specific command is often :OUTPut:PROTection:CLEar
            self.instrument.write('*CLS')
            # For specific instruments, you might add: self.instrument.write(':OUTPut:PROTection:CLEar')
            print("PSU protection trip state cleared (*CLS sent).")
        except Exception as e:
            print(f"Error clearing protection state: {e}")
            
    # -----------------------
    # Spectrum Analyzer (SA) / VNA Commands
    # -----------------------
    def do_rf_set_center_freq(self, frequency):
        """
        Sets the center frequency of the measurement sweep.

        Uses the common ':SENSe:FREQuency:CENTer' command.

        :param frequency: The desired center frequency (e.g., "1GHz" or "1000000000").
        :type frequency: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: rf_set_center_freq <frequency>
        Example: rf_set_center_freq 2.45GHz
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not frequency:
            print("Usage: rf_set_center_freq <frequency>")
            return
        try:
            self.instrument.write(f':SENSe:FREQuency:CENTer {frequency}')
            print(f"Center frequency set to: {frequency}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting center frequency: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_rf_set_span(self, span):
        """
        Sets the frequency span (range) of the measurement sweep.

        Uses the common ':SENSe:FREQuency:SPAN' command.

        :param span: The desired frequency span (e.g., "10MHz").
        :type span: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: rf_set_span <span_frequency>
        Example: rf_set_span 10MHz
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not span:
            print("Usage: rf_set_span <span_frequency>")
            return
        try:
            self.instrument.write(f':SENSe:FREQuency:SPAN {span}')
            print(f"Span set to: {span}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting span: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_rf_set_power(self, power_level):
        """
        Sets the reference power level (Ref Level for SA) or the source power (VNA).

        Uses the common commands: ':DISPlay:WINDow:TRACe:Y:RLEVel' (SA) or ':SOURce:POWer:LEVel' (VNA).
        The input power can include engineering suffixes (e.g., 'dBm').

        :param power_level: The desired power level (e.g., "-20dBm").
        :type power_level: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: rf_set_power <power_level_dBm>
        Example: rf_set_power -10dBm
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not power_level:
            print("Usage: rf_set_power <power_level>")
            return
        try:
            # VNA/Signal Generator standard:
            self.instrument.write(f':SOURce:POWer:LEVel {power_level}')
            # SA standard (often simplified via DISPlay) - users can use 'write' if this fails:
            # self.instrument.write(f':DISPlay:WINDow1:TRACe1:Y:RLEVel {power_level}')
            print(f"RF Power/Ref Level set to: {power_level}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting RF power level: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")

    def do_sa_set_rbw_vbw(self, bandwidths):
        """
        Sets the Resolution Bandwidth (RBW) and Video Bandwidth (VBW).

        These are critical settings for noise floor and trace averaging on SAs.

        Uses the common ':SENSe:BANDwidth:RESolution' and ':SENSe:BANDwidth:VIDeo' commands.

        :param bandwidths: Comma-separated string of RBW and VBW (e.g., "10kHz,3kHz").
        :type bandwidths: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects any command.

        Usage: sa_set_rbw_vbw <rbw>,<vbw>
        Example: sa_set_rbw_vbw 10kHz,3kHz
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not bandwidths:
            print("Usage: sa_set_rbw_vbw <rbw>,<vbw>")
            return
        try:
            rbw, vbw = map(str.strip, bandwidths.split(','))
            self.instrument.write(f':SENSe:BANDwidth:RESolution {rbw}')
            self.instrument.write(f':SENSe:BANDwidth:VIDeo {vbw}')
            print(f"RBW set to {rbw}, VBW set to {vbw}")
        except ValueError:
            print("Invalid format. Use: <rbw>,<vbw> (e.g., 10kHz,3kHz)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting bandwidths: {e}")
        except Exception as e:
            print(f"Unexpected error: {e}")

    def do_sa_read_marker(self, marker_number):
        """
        Queries the frequency and amplitude of a specified marker on the current trace.

        This usually assumes the marker is already active and placed (e.g., on a peak).

        Uses the common ':CALCulate:MARKer<n>:X?' and ':CALCulate:MARKer<n>:Y?' commands.

        :param marker_number: The marker number (typically 1 to 4).
        :type marker_number: str

        :raises pyvisa.errors.VisaIOError: If the query fails.

        Usage: sa_read_marker <marker_number>
        Example: sa_read_marker 1
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not marker_number:
            print("Usage: sa_read_marker <marker_number>")
            return
        try:
            mkr = marker_number.strip()
            freq = self.instrument.query(f':CALCulate:MARKer{mkr}:X?')
            amp = self.instrument.query(f':CALCulate:MARKer{mkr}:Y?')
            print(f"Marker {mkr}: Frequency = {freq.strip()} Hz, Amplitude = {amp.strip()} dBm")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error reading marker: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_vna_set_sweep(self, start_stop_points):
        """
        Sets the VNA sweep start frequency, stop frequency, and number of points.

        Uses the common ':SENSe:FREQuency:STARt', ':SENSe:FREQuency:STOP', and ':SENSe:SWEep:POINts' commands.

        :param start_stop_points: Comma-separated string of start_freq, stop_freq, and sweep_points (e.g., "1GHz,2GHz,201").
        :type start_stop_points: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects any command.

        Usage: vna_set_sweep <start_freq>,<stop_freq>,<points>
        Example: vna_set_sweep 1GHz,2GHz,201
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not start_stop_points:
            print("Usage: vna_set_sweep <start_freq>,<stop_freq>,<points>")
            return
        try:
            start, stop, points = map(str.strip, start_stop_points.split(','))
            self.instrument.write(f':SENSe:FREQuency:STARt {start}')
            self.instrument.write(f':SENSe:FREQuency:STOP {stop}')
            self.instrument.write(f':SENSe:SWEep:POINts {points}')
            print(f"VNA sweep set: {start} to {stop}, {points} points.")
        except ValueError:
            print("Invalid format. Use: <start_freq>,<stop_freq>,<points> (e.g., 1GHz,2GHz,201)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting VNA sweep parameters: {e}")
        except Exception as e:
            print(f"Unexpected error: {e}")

    def do_vna_measure_sparam(self, sparam_format):
        """
        Sets the active S-parameter measurement and format.

        Uses the common ':CALCulate:PARameter:DEFine' and ':CALCulate:FORMat' commands.

        :param sparam_format: Comma-separated string of S-parameter (Smn) and display format (e.g., "S21,MLOG").
            Common Formats: MLOG (Log Magnitude), PHAS (Phase), SWR (VSWR).
        :type sparam_format: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects any command.

        Usage: vna_measure_sparam <sparam>,<format>
        Example: vna_measure_sparam S21,MLOG
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not sparam_format:
            print("Usage: vna_measure_sparam <sparam>,<format>")
            return
        try:
            sparam, dformat = map(str.strip, sparam_format.split(','))
            sparam_scpi = sparam.upper()
            dformat_scpi = dformat.upper()

            # 1. Define the S-parameter measurement (Trace 1 is often the default/only trace)
            self.instrument.write(f':CALCulate:PARameter:DEFine "Trc1_{sparam_scpi}",{sparam_scpi}')
            # 2. Assign the format
            self.instrument.write(f':CALCulate:SELected:FORMat {dformat_scpi}')

            print(f"VNA Measurement set to {sparam_scpi} with format {dformat_scpi}.")
        except ValueError:
            print("Invalid format. Use: <sparam>,<format> (e.g., S21,MLOG)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting S-parameter measurement: {e}")
        except Exception as e:
            print(f"Unexpected error: {e}")
            
    def do_vna_set_trace(self, trace_param_window):
        """
        Sets the measurement parameter (Smn) and format for a specific trace and window.

        VNA Trace commands are typically structured around the window, trace, and parameter name.
        Uses ':CALCulate:PARameter:SELect' and ':DISPlay:WINDow:TRACe:FEED'.

        :param trace_param_window: Comma-separated string of trace, S-parameter (e.g., S11), and window (e.g., "2,S21,1").
        :type trace_param_window: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects any command.

        Usage: vna_set_trace <trace_num>,<sparam>,<window_num>
        Example: vna_set_trace 2,S21,1 (Sets Trace 2 in Window 1 to show S21)
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not trace_param_window:
            print("Usage: vna_set_trace <trace_num>,<sparam>,<window_num>")
            return
        try:
            tnum, sparam, wnum = map(str.strip, trace_param_window.split(','))
            sparam_scpi = sparam.upper()

            # 1. Select/Define the parameter name associated with the trace
            # VNA's often link a defined parameter to a trace number.
            self.instrument.write(f':CALCulate{wnum}:PARameter{tnum}:DEFine "{sparam_scpi}"')
            # 2. Assign the trace to display the newly defined parameter
            self.instrument.write(f':DISPlay:WINDow{wnum}:TRACe{tnum}:FEED "{sparam_scpi}"')

            print(f"Trace {tnum} in Window {wnum} now displays {sparam_scpi}.")
        except ValueError:
            print("Invalid format. Use: <trace_num>,<sparam>,<window_num> (e.g., 2,S21,1)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting VNA trace: {e}")
        except Exception as e:
            print(f"Unexpected error: {e}")

    def do_vna_query_data(self, arg):
        """
        Queries the measured trace data (the current S-parameter trace).

        Uses the common ':CALCulate:SELected:DATA:FDATa?' command to retrieve formatted data.

        :raises pyvisa.errors.VisaIOError: If the query fails or times out.

        Usage: vna_query_data
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            # Query the formatted data (FDATa) which is the currently displayed trace
            data = self.instrument.query(':CALCulate:SELected:DATA:FDATa?')
            # Data returned is a string of comma-separated values (CSV).
            # The length can be very long (e.g., 801 points * 2 values per point)
            data_points = data.strip().split(',')
            print(f"VNA Trace Data Queried: {len(data_points)} values received.")
            print(f"First 10 values: {data_points[:10]}...")
            print("\nNote: Use this with Python scripting for saving/plotting.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error querying VNA data: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_rf_screen_capture(self, filename_format):
        """
        Captures the instrument's screen and saves the image data to a local file (PNG or JPEG).

        Uses common SCPI commands like ':DISPlay:DATA?'. The output format is determined
        by the file extension and passed to the instrument.

        :param filename_format: The local file path and requested format, e.g., "capture.png".
        :type filename_format: str

        :raises pyvisa.errors.VisaIOError: If the instrument communication fails.
        :raises Exception: If the file cannot be written.

        Usage: rf_screen_capture <filename.png>
        Example: rf_screen_capture rf_test_results.jpeg
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not filename_format:
            print("Usage: rf_screen_capture <filename.png|filename.jpeg>")
            return
        
        try:
            filename = filename_format.strip()
            extension = filename.split('.')[-1].upper()

            if extension in ["PNG", "JPEG", "JPG"]:
                scpi_format = extension
                # Keysight/Agilent standard: Request the specific format
                SCPI_COMMAND = f':DISPlay:DATA? {scpi_format}' 
                
                print(f"Requesting screen data ({scpi_format})...")

                # PyVISA binary transfer: Read the data block.
                image_data = self.instrument.query_binary_values(
                    SCPI_COMMAND, 
                    datatype='s',
                    container=bytes,
                    is_termination_char=False
                )
                
                if not image_data:
                    print("Error: Received no image data from the instrument.")
                    return

                with open(filename, 'wb') as f:
                    f.write(image_data)
                
                print(f"Screen capture successfully saved to: {filename}")
            else:
                print("Unsupported format for screen capture. Use PNG or JPEG/JPG.")

        except pyvisa.errors.VisaIOError as e:
            print(f"Error communicating with instrument: {e}")
            print("HINT: Check instrument manual for the exact SCPI screen capture command.")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    # -----------------------
    # Load Setter (e.g., Electronic Load) Commands
    # -----------------------
    
    def do_eload_set_mode(self, mode):
        """
        Sets the electronic load's operating mode.

        Uses the common ':FUNCtion:MODE' command.

        :param mode: The desired operating mode ("CURR" for CC, "VOLT" for CV, "RES" for CR, or "POW" for CP).
        :type mode: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: eload_set_mode <CURR|VOLT|RES|POW>
        Example: eload_set_mode CURR
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not mode:
            print("Usage: eload_set_mode <CURR|VOLT|RES|POW>")
            return
        try:
            scpi_mode = mode.strip().upper()
            if scpi_mode not in ["CURR", "VOLT", "RES", "POW"]:
                print("Invalid mode. Use CURR, VOLT, RES, or POW.")
                return
            self.instrument.write(f':FUNCtion:MODE {scpi_mode}')
            print(f"Electronic Load mode set to: {scpi_mode}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting load mode: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_eload_set_current(self, current):
        """
        Sets the programmed current value for Constant Current (CC) mode.

        Uses the common ':CURRent' command.

        :param current: The desired current draw value (float, in Amps).
        :type current: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: eload_set_current <amps>
        Example: eload_set_current 1.5
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not current:
            print("Usage: eload_set_current <amps>")
            return
        try:
            self.instrument.write(f':CURRent {current}')
            print(f"CC current set to {current} A.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting current: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_eload_set_voltage(self, voltage):
        """
        Sets the programmed voltage value for Constant Voltage (CV) mode.

        Uses the common ':VOLTage' command.

        :param voltage: The desired voltage value (float, in Volts).
        :type voltage: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: eload_set_voltage <volts>
        Example: eload_set_voltage 5.0
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not voltage:
            print("Usage: eload_set_voltage <volts>")
            return
        try:
            self.instrument.write(f':VOLTage {voltage}')
            print(f"CV voltage set to {voltage} V.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting voltage: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_eload_set_resistance(self, resistance):
        """
        Sets the programmed resistance value for Constant Resistance (CR) mode.

        Uses the common ':RESistance' command.

        :param resistance: The desired resistance value (float, in Ohms).
        :type resistance: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: eload_set_resistance <ohms>
        Example: eload_set_resistance 10.5
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not resistance:
            print("Usage: eload_set_resistance <ohms>")
            return
        try:
            self.instrument.write(f':RESistance {resistance}')
            print(f"CR resistance set to {resistance} Ohm.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting resistance: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_eload_set_power(self, power):
        """
        Sets the programmed power value for Constant Power (CP) mode.

        Uses the common ':POWer' command.

        :param power: The desired power value (float, in Watts).
        :type power: str

        :raises pyvisa.errors.VisaIOError: If the instrument rejects the command.

        Usage: eload_set_power <watts>
        Example: eload_set_power 50
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not power:
            print("Usage: eload_set_power <watts>")
            return
        try:
            self.instrument.write(f':POWer {power}')
            print(f"CP power set to {power} W.")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting power: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
    
    def do_eload_input_on(self, arg):
        """
        Turns the electronic load input ON (starts drawing current).

        Uses the common ':INPut ON' command.

        :raises Exception: If command fails.

        Usage: eload_input_on
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            self.instrument.write(':INPut ON')
            print("Electronic Load Input ON (Load active).")
        except Exception as e:
            print(f"Error enabling load input: {e}")

    def do_eload_input_off(self, arg):
        """
        Turns the electronic load input OFF (stops drawing current).

        Uses the common ':INPut OFF' command.

        :raises Exception: If command fails.

        Usage: eload_input_off
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            self.instrument.write(':INPut OFF')
            print("Electronic Load Input OFF (Load inactive).")
        except Exception as e:
            print(f"Error disabling load input: {e}")
            
    def do_eload_measure_input(self, measure_type):
        """
        Measures the actual input voltage or current being sunk by the load.

        Uses the common ':MEASure:<TYPE>?' command.

        :param measure_type: The type of measurement to take ("VOLT" or "CURR").
        :type measure_type: str

        :raises Exception: If command fails.

        Usage: eload_measure_input <VOLT|CURR>
        Example: eload_measure_input VOLT
        """
        if not self.instrument:
            print("No device selected.")
            return
        measure_type = measure_type.strip().upper()
        if measure_type not in ["VOLT", "CURR"]:
            print("Usage: eload_measure_input <VOLT|CURR>")
            return
        try:
            measurement = self.instrument.query(f':MEASure:{measure_type}?')
            unit = 'V' if measure_type == 'VOLT' else 'A'
            print(f"ELoad Input {measure_type}: {measurement.strip()} {unit}")
        except Exception as e:
            print(f"Error measuring ELoad input {measure_type}: {e}")
            
    def do_eload_set_slew(self, rate):
        """
        Sets the Constant Current (CC) mode slew rate (A/µs or A/ms).

        This defines how quickly the load can transition between two current levels,
        which is essential for simulating sudden changes in a power supply's demands.
        The rate is typically in A/s.

        Uses the common ':CURRent:SLEW:RATE' command.

        :param rate: The desired current slew rate (float, in Amperes/second).
        :type rate: str

        :raises Exception: If command fails.

        Usage: eload_set_slew <amps_per_second>
        Example: eload_set_slew 0.1 (0.1 A/s)
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not rate:
            print("Usage: eload_set_slew <amps_per_second>")
            return
        try:
            self.instrument.write(f':CURRent:SLEW:RATE {rate}')
            print(f"CC Slew Rate set to {rate} A/s.")
        except Exception as e:
            print(f"Error setting current slew rate: {e}")
            
    def do_eload_set_transient(self, levels_timing):
        """
        Configures the Electronic Load for Constant Current (CC) transient operation.

        Sets two current levels (A and B) and the timing (e.g., rise/fall time, pulse width).
        The load must typically be set to CC mode first.

        Uses common commands like ':CURRent:STATic', ':CURRent:TRANsient:LEVel', and ':CURRent:TRANsient:TIME'.

        :param levels_timing: Comma-separated string of Level A, Level B, and Pulse Width (e.g., "0.5A,2.0A,10ms").
        :type levels_timing: str

        :raises ValueError: If the input format is incorrect.
        :raises pyvisa.errors.VisaIOError: If the instrument rejects any command.

        Usage: eload_set_transient <level_A>,<level_B>,<pulse_width>
        Example: eload_set_transient 0.5A,2.0A,10ms
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not levels_timing:
            print("Usage: eload_set_transient <level_A>,<level_B>,<pulse_width>")
            return
        try:
            level_a, level_b, pulse_width = map(str.strip, levels_timing.split(','))

            # Set the static/base current (Level A)
            self.instrument.write(f':CURRent:STATic {level_a}')
            # Set the transient current (Level B)
            self.instrument.write(f':CURRent:TRANsient:LEVel {level_b}')
            # Set the pulse width/time (SCPI for this varies, using a common one)
            self.instrument.write(f':CURRent:TRANsient:PULSe:WIDTh {pulse_width}')
            # Set the function to transient mode (important final step)
            self.instrument.write(':FUNCtion:MODE TRANsient')

            print(f"ELoad set for transient test: A={level_a}, B={level_b}, Width={pulse_width}.")
        except ValueError:
            print("Invalid format. Use: <level_A>,<level_B>,<pulse_width> (e.g., 0.5A,2.0A,10ms)")
        except pyvisa.errors.VisaIOError as e:
            print(f"Error setting transient mode: {e}")
        except Exception as e:
            print(f"An unexpected error occurred: {e}")
            
    def do_eload_set_ovl(self, voltage):
        """
        Sets the Over-Voltage Limit (OVL) protection threshold for the Electronic Load.

        If the input voltage exceeds this value, the load will trip and turn OFF the input.
        Uses the common ':VOLTage:PROTection:LEVel' command.

        :param voltage: The OVL limit voltage (float).
        :type voltage: str

        :raises Exception: If command fails.

        Usage: eload_set_ovl <voltage>
        Example: eload_set_ovl 6.0
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not voltage:
            print("Usage: eload_set_ovl <voltage>")
            return
        try:
            self.instrument.write(f':VOLTage:PROTection:LEVel {voltage}')
            print(f"OVL set to {voltage} V")
        except Exception as e:
            print(f"Error setting OVL: {e}")
            
    def do_eload_set_opl(self, power):
        """
        Sets the Over-Power Limit (OPL) protection threshold for the Electronic Load.

        If the input power (V * I) exceeds this value, the load will trip and turn OFF the input.
        Uses the common ':POWer:PROTection:LEVel' command.

        :param power: The OPL limit power (float, in Watts).
        :type power: str

        :raises Exception: If command fails.

        Usage: eload_set_opl <power_watts>
        Example: eload_set_opl 150
        """
        if not self.instrument:
            print("No device selected.")
            return
        if not power:
            print("Usage: eload_set_opl <power_watts>")
            return
        try:
            self.instrument.write(f':POWer:PROTection:LEVel {power}')
            print(f"OPL set to {power} W")
        except Exception as e:
            print(f"Error setting OPL: {e}")

    # -----------------------
    # Diagnostics & Robustness
    # -----------------------
    def do_ping_device(self, arg):
        """
        Pings the connected instrument to verify communication (uses *OPC?).

        :raises pyvisa.errors.VisaIOError: If the query fails or times out.

        Usage: ping_device
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            # Some instruments respond 1 to *OPC? when operations complete
            resp = self.instrument.query('*OPC?').strip()
            if resp == '1' or resp.lower() == 'true':
                print("Instrument communication OK.")
            else:
                print(f"Unexpected response: {resp}")
        except pyvisa.errors.VisaIOError as e:
            print(f"Device ping failed (VisaIOError): {e}")
        except Exception as e:
            print(f"Device ping failed: {e}")

    def do_check_capabilities(self, arg):
        """
        Checks instrument capabilities (queries *IDN?, *OPT? and tries some SCPI probes).

        Performs non-exhaustive attempts to query common SCPI commands to give
        an indication of the instrument's supported features (DMM, Scope, etc.).

        :raises Exception: For general errors during the capability check process.

        Usage: check_capabilities
        """
        if not self.instrument:
            print("No device selected.")
            return
        try:
            print("Checking instrument capabilities...\n")
            try:
                idn = self.instrument.query('*IDN?').strip()
                print(f"IDN: {idn}")
            except Exception:
                print("IDN query not supported or failed.")

            try:
                opts = self.instrument.query('*OPT?').strip()
                print(f"Options: {opts}")
            except Exception:
                print("Options query not supported.")

            # Probe a few common SCPI commands and collect those that respond without exception
            candidate_cmds = [
                ':MEASure:VOLTage:DC?',
                ':MEASure:VOLTage:AC?',
                ':MEASure:CURRent:DC?',
                ':MEASure:RESistance?',
                ':WAVeform:DATA?',
                ':TRIGger:SOURce?',
                ':OUTPut:STATe?'
            ]
            supported = []
            for cmd in candidate_cmds:
                try:
                    # Use query but allow that commands which require arguments may error; ignore their error
                    _ = self.instrument.query(cmd)
                    supported.append(cmd)
                except Exception:
                    # ignore
                    pass

            if supported:
                print("\nLikely supported SCPI commands (best-effort probe):")
                for c in supported:
                    print(f"  {c}")
            else:
                print("\nCould not auto-detect SCPI feature set from probes.")
        except Exception as e:
            print(f"Error checking capabilities: {e}")

    # -----------------------
    # Misc / Exit
    # -----------------------
    def do_exit(self, arg):
        """
        Exits the console and closes the instrument connection.

        Saves command history before exiting.

        Usage: exit
        """
        # Save history
        try:
            readline.write_history_file(HISTORY_FILE)
        except Exception:
            pass

        if self.instrument:
            try:
                print(f"Closing connection to {self.selected_device_id}...")
                self.instrument.close()
            except Exception:
                pass
        print("Exiting console. Goodbye!")
        sys.exit(0)

    # Support Control-D to exit
    def do_EOF(self, arg):
        """Handles EOF (Ctrl+D) as an exit command."""
        print()
        return self.do_exit(arg)

    def postloop(self):
        """Cleanup function called after the command loop finishes."""
        # Cleanup on loop exit (if any)
        if self.instrument:
            try:
                self.instrument.close()
            except Exception:
                pass
        try:
            readline.write_history_file(HISTORY_FILE)
        except Exception:
            pass

# -----------------------
# Main
# -----------------------
if __name__ == '__main__':
    try:
        Console().cmdloop()
    except KeyboardInterrupt:
        print("\nCtrl+C detected. Exiting console. Goodbye!")
    except Exception as e:
        print(f"A fatal error occurred: {e}")
