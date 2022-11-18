
def test_method(test_name):
    methods={
        "Test A:Cold":[["IEC 60068-2-1","2007-03"],["low temperature operation","low temperature storage","low temperature operation","lts","cold temperature","Cold Storage","CL/04","Cold Operation","CL/09"]],
        "Test B: Dry Heat": [["IEC 60068-2-2", "2007-07"],["high temperature operation","high temperature storage","high temperature operation","lts","warm temperature","Warm Storage","CL/03","Warm Operation","CL/08"]],
        "Test Na: Rapid change of temperature with prescribed time of transfer": [["IEC 60068-2-14", "2009-01"],['RAPID CHANGE OF TEMPERATURE WITH SPECIFIED TRANSITION DURATION','Thermal Shock','CL/01','Thermal Shocks Pre-Aging','CL/02']],
        "Test Nb: Change of temperature with specified rate of change": [["IEC 60068-2-14", "2009-01"],['TEMPERATURE CYCLING WITH SPECIFIED CHANGE RATE','TEMPERATURE CYCLING WITH SPECIFIED TEMPERATURE CHANGE RATE','LT/01','Thermal Cycling']],
        "Temperature Step test": [["ISO 16750 – 4", "2010-04"],['TEMPERATURE STEPS TEST','TEMPERATURE STEP test','Temperature range step','Temperature range','TEMPERATURE STEP','Step Temperature Test','TEMPERATURE STEPS','CL/07']],
        "Ice water shock test": [["ISO 16750 – 4", "2010-04"],['Ice water shock','IWS']],
        "Test Db: Damp heat, cyclic (12 h + 12 h cycle)":[["IEC 60068-2-30","2005-08"],['0']],
        "Test Z/AD: Composite temperature/humidity cyclic test": [["IEC 60068-2-38", "2021-03"],['COMPOSITE TEMPERATURE/HUMIDITY CYCLIC TEST','COMPOSITE TEMPERATURE/ HUMIDITY CYCLE TEST','Climatic sequence','CL/06']],
        "Test Z/AM: Combined cold/low air pressure tests": [["IEC 60068-2-40", "1976-01"],['CL/10', 'COLD AND LOW PRESSURE STORAGE']],
        "Test Cy: Damp heat, steady state": [["IEC 60068-2-67", "2019-07"],['0']],
        "Test Cab: Damp heat, steady state": [["IEC 60068-2-78", "2012-10"],['DAMP HEAT, STEADY STATE']],
        "Degrees of protection provided by enclosures against foreign objects and dust (IP Code)": [[
            "ISO 20653 \n IEC 60529", "2013-02 \n 2013-08"],['PROTECTION AGAINST FOREIGN OBJECTS',"dust",'PROTECTION AGAINST ACCESS','PROTECTION AGAINST FOREIGN OBJECTS TEST']],
        "Degrees of protection against water (IP Code)": [["ISO 20653 \n IEC 60529" "2013-02 \n 2013-08"],['WATER INTRUSION','HIGH-PRESSURE TEST','IPX9K','IPX6K','WATER WADING','IPX7','IP X2']],
        "Test Ka: Salt mist": [["IEC 60068-2-11","2021-03"],['SALT SPRAY, LEAKAGE AND FUNCTION']],
        "Test Kb: Salt mist, cyclic (sodium chloride solution)": [["IEC 60068-2-52","2017-11"],['0']],
        "Chemical loads": [["ISO 16750-5", "2010-04"],['CHEMICAL LOADS']],
        "Test Fc:Vibration (Sinusoidal)":[["ISO 16750-3 \n IEC 60068-2-6",	"2012-12 \n 2007-12"],['Resonance Point','VI/05','Sinusoidal vibrations','Sinus vibrations']],
        "Test Fh: Vibration, broadband random and guidance": [["ISO 16750-3 \n IEC 60068-2-64", "2012-12 \n 2019-10"],
                                                              ['RANDOM VIBRATION','VI/07']],
        "Test Fi: Vibration – Mixed mode": [["ISO 16750-3 \n IEC 60068-2-80", "2012-12 \n 2005-05"], ['0']],
        "Test Ea and guidance: Shock": [["ISO 16750- \n IEC 60068-2-27", "2012-12 \n 2008-02"], ['MECHANICAL SHOCK','Mounting operation shock','MS/02','Collision Impact','MS/03']],
        "Test Ec: Free fall (Procedure 1)": [["ISO 16750-3 \n IEC 60068-2-31", "2012-12 \n 2008-05"], ['FREE FALL','drop','MS/01']],
        "Insulation resistance": [["ISO 16750-2", "2012-11"], ['0']],
        "Withstand voltage": [["ISO 16750-2", "2012-11"], ['0']],

    }
    for key in methods:
        for name in methods[key][1]:
            if name.lower().replace(" ", "") in test_name.lower().replace(" ", ""):
                # print(key)
                return key


