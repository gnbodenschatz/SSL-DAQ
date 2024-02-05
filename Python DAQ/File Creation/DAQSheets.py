from datetime import date
import xlsxwriter

class DAQSheets:
    def __init__(self, path, name, specs):
        self.date = date.today()
        self.fileName = name
        if path is not None:
            self.path = path
        else:
            self.path = './FileCreation/Files'
        self.SweepTime = specs["Sweep Time"]
        print(f'Sweep time is {self.SweepTime}')



    def createSheetDefault(self):
        wb = xlsxwriter.Workbook(f'{self.path}{self.date}{self.fileName}Combined.xlsx')
        data = wb.add_worksheet()
        print('Worksheet made\n')

        #Write Specs
        data.write(0,0,'Date:')
        data.write(2,0,'Sample:')
        data.write(4,0,'DAS Version Num')
        data.write(5,0,'Scans')
        data.write(6,0,'Sweep Time (s)')
        data.write(6,1,self.SweepTime)
        data.write(7,0,'Center Field (G)')
        data.write(8,0,'Sweep Width (G)')
        data.write(9,0,'Mod Freq (Hz)')
        data.write(10,0,'Mod Ampl (G)')
        data.write(11,0,'uW Freq (GHz)')
        data.write(12,0,'uW Power (mW)')
        data.write(13,0,'Gate Bias (V)')
        data.write(14,0,'Diode Bias (V)')
        data.write(15,0,'DC Current (A)')
        data.write(16,0,'PA Sens (A/V)')
        data.write(17,0,'Phase (deg)')
        data.write(18,0,'LIA Sens (A/V)')
        data.write(19,0,'Time Const (s)')
        data.write(20,0,'Angle (deg)')
        data.write(21,0,'Temperature (K)')
        data.write(22,0,'Harmonic')
        data.write(23,0,'Units')
        data.write(24,0,'Measured Current')

        #Create Column headers
        data.write(26,0,'Field')
        data.write(26,1,'avg_i')
        data.write(26,2,'avg_q')
        data.write(26,3,'avg_if')
        data.write(26,4,'avg_qf')
        data.write(26,5,'ind_i')
        data.write(26,6,'ind_q')
        data.write(26,7,'ind_if')
        data.write(26,8,'ind_qf')
        data.write(26,9,'der_i')
        data.write(26,10,'int_i')
        data.write(26,11,'der_q')
        data.write(26,12,'int_q')

        wb.close()
        print('Workbook closed\n')
        
    def createSheetSimple(self):
        wb1 = xlsxwriter.Workbook(f'{self.path}{self.date}{self.fileName}Split.xlsx')
        data = wb1.add_worksheet()
        specs = wb1.add_worksheet()

        #Write Specs
        specs.write(0,0,'Date:')
        specs.write(2,0,'Sample:')
        specs.write(4,0,'DAS Version Num')
        specs.write(5,0,'Scans')
        specs.write(6,0,'Sweep Time (s)')
        specs.write(7,0,'Center Field (G)')
        specs.write(8,0,'Sweep Width (G)')
        specs.write(9,0,'Mod Freq (Hz)')
        specs.write(10,0,'Mod Ampl (G)')
        specs.write(11,0,'uW Freq (GHz)')
        specs.write(12,0,'uW Power (mW)')
        specs.write(13,0,'Gate Bias (V)')
        specs.write(14,0,'Diode Bias (V)')
        specs.write(15,0,'DC Current (A)')
        specs.write(16,0,'PA Sens (A/V)')
        specs.write(17,0,'Phase (deg)')
        specs.write(18,0,'LIA Sens (A/V)')
        specs.write(19,0,'Time Const (s)')
        specs.write(20,0,'Angle (deg)')
        specs.write(21,0,'Temperature (K)')
        specs.write(22,0,'Harmonic')
        specs.write(23,0,'Units')
        specs.write(24,0,'Measured Current')

        #Create Column headers
        data.write(0,0,'Field')
        data.write(0,1,'avg_i')
        data.write(0,2,'avg_q')
        data.write(0,3,'avg_if')
        data.write(0,4,'avg_qf')
        data.write(0,5,'ind_i')
        data.write(0,6,'ind_q')
        data.write(0,7,'ind_if')
        data.write(0,8,'ind_qf')
        data.write(0,9,'der_i')
        data.write(0,10,'int_i')
        data.write(0,11,'der_q')
        data.write(0,12,'int_q')
        
        wb1.close()


specs = {
    "Sample" : 'This is the device',
    "Sweep Time" : '100s',
    "Center Field" : '3390',
    "Sweep Width" : '400',
    "Mod Freq" : '1000',
    "uW Freq" : '9.3',
    "uw Power" : '',
    "Gate Bias" : '',
    "Diode Bias" : '',
    "DC Current" : '',
    "PA Sens" : '',
    "Phase" : '',
    "LIA Sens" : '',
    "Time Const" : '',
    "Angle" : '',
    "Temp" : '',
    "Harmonic" : '',
    "Units" : '',
    "Measured Current" : ''
}
sheet = DAQSheets('./Python DAQ/File Creation/Files/',specs)
sheet.createSheetDefault()
sheet.createSheetSimple()

