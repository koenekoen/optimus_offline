"""
|-------------------------------------|
|  Supply Chain Optimisation Toolbox  |
|-------------------------------------|

New in v0.9.0.0:
    Automated Analyses windows now resets correctly after trade-off analyses
    Food Group Exclusions are now loaded correctly from csv
    Added automated analysis: Increase Prices (for loc/reg purchases)
    Added option to supply tactical demand (shopping list) - filterable
    Improved constraint handling for periods that fall outside the current time horizon

"""

# Import external packages
from Tkinter import * # Allows creation of the GUI
import tkMessageBox # Extension of Tkinter (pop-ups)
import ttk # Aesthetic upgrade of Tkinter
from pulp import * # Interface to COIN-OR solver
import csv # Allows manipulation of excel files
import time # Allows the tracking of time
import os # Allows access to (sub)folders
import pdb # Allows tracing/breakpoints using pdb.set_trace()
import shutil # Allows the copying of files
import pickle # Allows the saving/loading of variables
import sys, traceback # Allows the tracing of errors using traceback.print_exc(file=sys.stdout)
from itertools import chain, combinations # Allows the creation of all subsets of a set
import datetime # Allows conversion of excel's ridiculous date format

class UNWFPModel:
    def __init__(self, root):
        '''
        This code runs automatically when the file is opened.
        It loads the data from UpdateValues.xlsm and creates the GUI.
        '''

        print "Welcome to AID-M: WFP's Assistant for Integrated Decision-Making"
        print " "

        # Load data
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'data')
        path = os.path.join(dest_dir, 'Check.csv')
        out = open(path,"rb")
        c = csv.reader(out, dialect='excel')
        check = next(c,None)[0]
        out.close()
        if check == "1":
            print "No change in data detected"
            print "Loading data from previous state..."
            self.load_quick()
        else:
            print "New data detected"
            print "Loading data from file..."
            self.load_thorough()
        print "Data loaded!"
        print " "

        # Create GUI
        print "Creating GUI..."
        print " "
        self.draw_GUI(root)
        print "Ready when you are!"
        print " "

        # Run procedures (testing / temporary analyses)
        #self.auto_temp()

    def load_thorough(self):
        '''
        Load, filter, and connect all the relevant data from the .csv files
        that were prepared through UpdateValues.xlsm.
        '''

        print "Loading nutritional data..."
        # FCS is predefined
        self.fcsgroups = ["Main staples","Pulses","Vegetables","Fruit","Meat and fish","Milk","Sugar","Oil","Condiments","Other"]
        self.weight = {}
        self.weight["Main staples"] = 2
        self.weight["Pulses"] = 3
        self.weight["Vegetables"] = 1
        self.weight["Fruit"] = 1
        self.weight["Meat and fish"] = 4
        self.weight["Milk"] = 4
        self.weight["Sugar"] = 0.5
        self.weight["Oil"] = 0.5
        self.weight["Condiments"] = 0
        self.weight["Other"] = 0 # to handle exceptions and SNFs
        # Grab nutritional data for each commodity from NutVal
        fileloc = os.path.dirname(os.path.abspath(__file__))
        dataloc = os.path.join(fileloc,'data')
        csvloc = os.path.join(dataloc,'Nutritional Values.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        # Define the sets
        self.supcom = [] # Commodity archetype
        self.commodities = [] # Specific commodity
        self.foodgroups = []
        self.nutrients = next(myreader,None)[2:-2]
        # Define the parameters
        self.nutval = {}
        self.group = {}
        self.sup = {}
        self.fcs = {}
        # Load the data
        for item in myreader:
            if item[0] != "" and item[0] != "0":
                self.supcom.append(item[0])
                self.commodities.append(item[1])
                self.foodgroups.append(item[13][1:-1])
                i=2
                for l in self.nutrients:
                    self.nutval[item[1],l]=float(item[i])
                    i+=1
                self.group[item[1]]=item[i][1:-1]
                self.sup[item[1]] = item[0]
                if str(item[i+1]) in self.fcsgroups:
                    self.fcs[item[1]]=item[i+1]
                else:
                    print "<<<Warning>>> " + item[i+1] + " is not an FCS food group"
                    self.fcs[item[1]]="Other"
        f.close()
        # Remove duplicates from sets
        self.commodities = list(set(self.commodities)) # By making a set of the food groups we remove double entries. Turning it into a list again allows for easier use of this set
        self.commodities.sort()
        self.supcom = list(set(self.supcom))
        self.supcom.sort()
        self.foodgroups = list(set(self.foodgroups))
        self.foodgroups.sort()
        # Grab the nutritional requirements for each beneficiary type
        csvloc = os.path.join(dataloc,'Nutritional Requirements.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        self.beneficiaries = []
        self.benlist = []
        self.nutreq = {}
        next(myreader,None)
        for item in myreader:
            self.beneficiaries.append(item[0])
            i=1
            v=0
            for l in self.nutrients:
                self.nutreq[item[0],l]=float(item[i])
                v+=float(item[i])
                i+=1
            if v>0:
                self.benlist.append(item[0])
        f.close()

        print "Loading beneficiary allocations..."
        csvloc = os.path.join(dataloc,'Beneficiary Allocations.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        self.FDPs = []
        self.periods = next(myreader,None)[3:]
        self.hc = {}
        self.dem = {}
        check = []
        for item in myreader:
            self.FDPs.append(item[1])
            self.hc[item[1]]=float(item[2])
            i=3
            v=0
            for t in self.periods:
                self.dem[item[0],item[1],t]=float(item[i])
                v+=float(item[i])
                if (item[0],t) in self.dem.keys(): # Also keep track of aggregate demand for each t
                    self.dem[item[0],t] += float(item[i])
                else:
                    self.dem[item[0],t] = float(item[i])
                i+=1
            if v>0:
                check.append(item[0])
        f.close()
        # Filter
        self.FDPs=list(set(self.FDPs))
        self.FDPs.sort()
        self.beneficiaries = [b for b in self.beneficiaries if b in check]
        self.benlist = [b for b in self.benlist if b in check]
        # NB: This list now contains activities for which we have demand and a nutritional profile, i.e. they can be optimised
        for b in self.beneficiaries:
            for i in self.FDPs:
                for t in self.periods:
                    if (b,i,t) not in self.dem.keys():
                        self.dem[b,i,t] = 0  # makes constraints easier to define

        print "Loading node capacities..."
        csvloc = os.path.join(dataloc,'Discharge Ports.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        self.DPs = []
        self.nodecap = {}
        self.sc = {}
        next(myreader,None)
        for item in myreader:
            self.DPs.append(item[0])
            self.hc[item[0]]=float(item[1])
            self.sc[item[0]]=float(item[2])
            i=3
            for t in self.periods:
                self.nodecap[item[0],t] = float(item[i])
                i+=1
        self.DPs=list(set(self.DPs))
        self.DPs.sort()
        f.close()
        csvloc = os.path.join(dataloc,'Extended Delivery Points.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        self.EDPs = []
        next(myreader,None)
        for item in myreader:
            self.EDPs.append(item[0])
            self.hc[item[0]]=float(item[1])
            self.sc[item[0]]=float(item[2])
            i=3
            for t in self.periods:
                self.nodecap[item[0],t] = float(item[i])
                i+=1
        self.EDPs=list(set(self.EDPs))
        self.EDPs.sort()
        f.close()

        print "Loading support costs..."
        csvloc = os.path.join(dataloc,'Support Costs.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.odocF = float(next(myreader,None)[1])
        self.odocCV = float(next(myreader,None)[1])
        self.dsc = float(next(myreader,None)[1])
        self.isc = float(next(myreader,None)[1])
        self.ltsh = float(next(myreader,None)[1])
        f.close()

        print "Loading upstream routes..."
        csvloc = os.path.join(dataloc,'SCIPS Routes.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.ISs = []
        self.missingcoms = []
        self.cost = {}
        self.dur = {}
        for item in myreader:
            if item[1] in self.DPs:
                if item[2] in self.commodities: # Note that we ignore connections that are not relevant to our CO to reduce complexity
                    self.ISs.append(item[0])
                    self.cost[item[0],item[1],item[2]]=float(item[3])
                    self.dur[item[0],item[1],item[2]]=float(item[4])
                else:
                    self.missingcoms.append(item[2])
        self.ISs = list(set(self.ISs))
        self.ISs.sort()
        f.close()
        # Overwrite lead times with historical values if not specified explicitly
        csvloc = os.path.join(dataloc,'Shipping Times.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        for item in myreader:
            if item[1] in self.DPs:
                for k in self.commodities:
                    if (item[0],item[1],k) in self.dur.keys():
                        if self.dur[item[0],item[1],k]==0:
                            self.dur[item[0],item[1],k]=float(item[2])
        # Add port processing times to shipping connections
        csvloc = os.path.join(dataloc,'Port Processing Times.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        for item in myreader:
            if item[0] in self.DPs:
                for lp in self.ISs:
                    for k in self.commodities:
                        if (lp,item[0],k) in self.dur.keys():
                            self.dur[lp,item[0],k]+=float(item[1])

        csvloc = os.path.join(dataloc,'Overland Routes.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.RSs = []
        for item in myreader:
            if item[1] in self.EDPs:
                self.RSs.append(item[0])
                for k in self.commodities:
                    self.cost[item[0],item[1],k] = float(item[2])
                    self.dur[item[0],item[1],k] = float(item[3])
        self.RSs = list(set(self.RSs))
        self.RSs.sort()
        f.close()
        for city in self.RSs:
            if city in self.ISs:
                self.ISs.remove(city) # RMs and LPs should not overlap; otherwise some procurement costs will be counted twice

        csvloc = os.path.join(dataloc,'Local Procurement Routes.csv')
        f = open(csvloc, "r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.LSs= []
        for item in myreader:
            if item[1] in self.EDPs:
                self.LSs.append(item[0])
                for k in self.commodities:
                    self.cost[item[0],item[1],k] = float(item[2])
                    self.dur[item[0],item[1],k] = float(item[3])
        self.LSs = list(set(self.LSs))
        self.LSs.sort()
        f.close()

        csvloc = os.path.join(dataloc,'C&V Routes.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.LMs = []
        for item in myreader:
            if item[1] in self.FDPs:
                self.LMs.append(item[0])
                for k in self.commodities:
                    self.cost[item[0],item[1],k] = float(item[2])
                    self.dur[item[0],item[1],k] = float(item[3])
        self.LMs = list(set(self.LMs))
        self.LMs.sort()
        f.close()

        print "Loading procurement options..."
        csvloc = os.path.join(dataloc,'SCIPS Prices.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        self.proccap = {}
        self.date = {}
        self.isGMO = {}
        self.countries = []
        self.incoterms = []
        self.sources = []
        next(myreader,None)
        for item in myreader: # item = (Origin country, named delivery place, incoterm, commodity, spec. commodity, gmo, packaging type, price, capacity (opt), processing time (opt))
            if item[1] in (self.ISs+self.DPs+self.RSs+self.LSs): # Discard procurement options for which no transportation link exists
                if float(item[7])>0: # Ignore entries for which the price is unknown
                    if item[4] in self.commodities:
                        src = item[0] + " - " + item[2]
                        self.sources.append(src) # The country+Incoterm identifier is used as a source
                        self.cost[src,item[1],item[4]] = float(item[7]) # Connects the named delivery place to a source
                        self.date[src,item[1],item[4]] = self.xldate2month(float(item[8]),0)
                        self.proccap[src, item[1],item[4]] = float(item[9])
                        self.dur[src, item[1],item[4]] = float(item[10])
                        if item[5]=="GMO":
                            self.isGMO[src, item[1],item[4]] = 1
                        else:
                            self.isGMO[src, item[1],item[4]] = 0
                        self.countries.append(item[0])
                        self.incoterms.append(item[2])
                    else:
                        self.missingcoms.append(item[4])
        f.close()
        self.countries = list(set(self.countries))
        self.countries.sort()
        self.incoterms = list(set(self.incoterms))
        self.incoterms.sort()
        self.sources = list(set(self.sources))
        self.sources.sort()

        csvloc = os.path.join(dataloc,'VAM Prices.csv')
        f = open(csvloc,"r")
        src = "Local Markets - C&V"
        self.sources.append(src) # Create a new source for C&V purchases
        myreader = csv.reader(f)
        next(myreader,None)
        for item in myreader:
            if item[1] in self.commodities:
                self.cost[src, item[0], item[1]] = float(item[2])
                self.proccap[src, item[0],item[1]] = float(item[3])
                self.dur[src, item[0],item[1]] = float(item[4])
                self.isGMO[src, item[0],item[1]] = float(item[5])
            else:
                self.missingcoms.append(item[1])
        f.close()

        print "Cross-referencing..."
        # Warning: Missing commodities
        self.missingcoms = list(set(self.missingcoms))
        self.missingcoms.sort()
        if self.missingcoms != []:
            print "<<<Warning>>> Not all commodities can be found in NutVal!"
            for item in self.missingcoms:
                print "  Missing: " + item
            print "If you want any of these commodities to be considered for purchase, add them to the NutVal tab in UpdateValues.xlsm"
            print " "
        # Filter self.cost (remove loose ends)
        disc_r = 0
        disc_p = 0
        self.avail = [] # The subset of commodities that is actually available for purchase
        temp=[]
        for key in self.cost.keys():
            if key[0] not in self.sources: # key = (ndp, location inside country, com)
                check = 0
                for src in self.sources:
                    if (src,key[0],key[2]) in self.cost.keys():
                        check = 1 # Commodity can be procured
                        self.avail.append(key[2])
                        break
                if check == 0 and key[0] not in self.DPs: # Routing option doesn't connect with a sourcing option
                    self.cost.pop(key,None) # Remove the option from consideration
                    disc_r += 1
            else: # key = (src, ndp, com)
                check = 0
                if key[1] in self.ISs:
                    for dp in self.DPs:
                        if (key[1],dp,key[2]) in self.cost.keys():
                            check = 1 # Commodity can be shipped
                            break
                if key[1] in self.RSs:
                    for i in (self.DPs+self.EDPs):
                        if (key[1],i,key[2]) in self.cost.keys():
                            check = 1 # Commodity can be shipped or transported overland
                            break
                if key[1] in self.LSs:
                    for edp in self.EDPs:
                        if (key[1],edp,key[2]) in self.cost.keys():
                            check = 1 # Commodity can be transported inland
                            break
                if key[1] in self.LMs:
                    for fdp in self.FDPs:
                        if (key[1],fdp,key[2]) in self.cost.keys():
                            check = 1 # Commodity can be transported inland
                            break
                if check == 0: # Procurement option doesn't connect with a routing option
                    self.cost.pop(key,None) # Remove the option from consideration
                    self.proccap.pop(key,None)
                    self.isGMO.pop(key,None)
                    disc_p +=1
                    temp.append(key)
        print "Removed " + str(disc_r) + " disconnected routes"
        print "Removed " + str(disc_p) + " disconnected procurement options"
        # Remove redundant commodities
        self.avail=list(set(self.avail))
        self.avail.sort()
        print "Removed " + str(len(self.commodities)-len(self.avail)) + " disconnected commodities"
        print " "
        self.commodities = self.avail
        for s in self.supcom:
            check = 0
            for k in self.commodities:
                if self.sup[k] == s:
                    check = 1
                    break
            if check == 0:
                self.supcom.remove(s)
        for key in self.cost.keys():
            if key[2] not in self.commodities:
                self.cost.pop(key,None)
        # Warn user of disconnected procurement options
        if len(temp) > 0:
            print "<<Warning>> AID-M removed "+str(len(temp))+" procurement options that did not connect to the CO's supply chain network:"
            temp = list(set(temp))
            temp.sort()
            for i in temp:
                print i
            print "If you want these procurement options to be considered, add a shipping or overland connection in UpdateValues.xlsm"
            print " "

        # Initialise default ration constraints
        self.user_add_com = {}
        for k in self.commodities:
            self.user_add_com[k]=[0,1000] # g/p/d of each commodity is between 0-1000g

        print "Loading downstream routes..."
        csvloc = os.path.join(dataloc,'DP2EDP Transport.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.arccap = {}
        for item in myreader:
            if item[0] in self.DPs and item[1] in self.EDPs:
                i=4
                for t in self.periods:
                    self.arccap[item[0],item[1],t] = float(item[i])
                    i+=1
                for k in self.commodities:
                    self.cost[item[0],item[1],k] = float(item[2])
                    self.dur[item[0],item[1],k] = float(item[3])
        f.close()
        csvloc = os.path.join(dataloc,'DP2DP Transport.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        for item in myreader:
            if item[0] in self.DPs and item[1] in self.DPs:
                i=4
                for t in self.periods:
                    self.arccap[item[0],item[1],t] = float(item[i])
                    i+=1
                for k in self.commodities:
                    self.cost[item[0],item[1],k] = float(item[2])
                    self.dur[item[0],item[1],k] = float(item[3])
        f.close()
        csvloc = os.path.join(dataloc,'EDP2FDP Transport.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        for item in myreader:
            if item[0] in self.EDPs and item[1] in self.FDPs:
                i=4
                for t in self.periods:
                    self.arccap[item[0],item[1],t] = float(item[i])
                    i+=1
                for k in self.commodities:
                    self.cost[item[0],item[1],k] = float(item[2])
                    self.dur[item[0],item[1],k] = float(item[3])
        f.close()
        csvloc = os.path.join(dataloc,'EDP2EDP Transport.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        for item in myreader:
            if item[0] in self.EDPs and item[1] in self.EDPs:
                i=4
                for t in self.periods:
                    self.arccap[item[0],item[1],t] = float(item[i])
                    i+=1
                for k in self.commodities:
                    self.cost[item[0],item[1],k] = float(item[2])
                    self.dur[item[0],item[1],k] = float(item[3])
        f.close()

        print "Loading initial inventories..."
        csvloc = os.path.join(dataloc,'Initial Inventory (DP).csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.inv = {}
        for i in (self.DPs+self.EDPs+self.ISs+self.LMs+self.LSs+self.RSs):
            for k in self.commodities:
                for t in self.periods:
                    self.inv[i,k,t] = 0 # Initialising initial inventory for each transshipment node makes the constraints easier to define
        for item in myreader:
            if item[0] in self.DPs:
                i=2
                for t in self.periods:
                    if item[i]=="":
                        self.inv[item[0],item[1],t]=0
                    else:
                        self.inv[item[0],item[1],t]=float(item[i])
                    i+=1
        f.close()
        csvloc = os.path.join(dataloc,'Initial Inventory (EDP).csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        for item in myreader:
            if item[0] in self.EDPs:
                i=2
                for t in self.periods:
                    if item[i]=="":
                        self.inv[item[0],item[1],t]=0
                    else:
                        self.inv[item[0],item[1],t]=float(item[i])
                    i+=1
        f.close()

        print "Loading additional demands..."
        csvloc = os.path.join(dataloc,'Activity Rations.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.baskets = {}
        self.feedingdays = {}
        for b in self.beneficiaries:
            for k in self.commodities:
                self.baskets[b,k] = 0 # Initialising for each node makes the constraints easier to define
                self.feedingdays[b,k] = 0
            self.baskets[b,"CASH"] = 0
            self.feedingdays[b,"CASH"] = 0
        for item in myreader:
            if item[0] in self.beneficiaries and item[1] in self.commodities:
                self.baskets[item[0],item[1]] = float(item[2])
                self.feedingdays[item[0],item[1]] = float(item[3])
            else:
                print " > Taxonomy not recognised: ", item
        f.close()
        for b in self.beneficiaries:
            if sum(self.baskets[b,k] for k in self.commodities) + self.baskets[b,"CASH"] == 0:
                print "<<<WARNING>>> No food basket defined for activity: " + b

        csvloc = os.path.join(dataloc,'Tactical Demand.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.tact_demand = {}
        self.tact_fdp = {}
        self.tact_com = {}
        self.tact_mon = {}
##        for i in self.FDPs:
##            for t in self.periods:
##                for k in self.commodities:
##                    self.tact_demand[i,k,t] = 0 # pre-loading allows for easier constraint definition
##                self.tact_demand[i,"CASH",t] = 0
        for item in myreader:
            i,k,t,d = item[0],item[1],self.xldate2month(float(item[2]),1),float(item[3])
            if i in self.FDPs and k in self.commodities and t in self.periods:
                self.tact_demand[i,k,t] = d
                if i in self.tact_fdp.keys():
                    self.tact_fdp[i] += d
                else:
                    self.tact_fdp[i] = d
                if k in self.tact_com.keys():
                    self.tact_com[k] += d
                else:
                    self.tact_com[k] = d
                if t in self.tact_mon.keys():
                    self.tact_mon[t] += d
                else:
                    self.tact_mon[t] = d
            else:
                print " > Taxonomy not recognised:  ",i,k,t,d
        f.close()

        print "Loading forecasts..."
        csvloc = os.path.join(dataloc,'Price Seasonality.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.fc_price={}
        for item in myreader:
            self.fc_price[item[0],item[1],"Jan"]=float(item[2])
            self.fc_price[item[0],item[1],"Feb"]=float(item[3])
            self.fc_price[item[0],item[1],"Mar"]=float(item[4])
            self.fc_price[item[0],item[1],"Apr"]=float(item[5])
            self.fc_price[item[0],item[1],"May"]=float(item[6])
            self.fc_price[item[0],item[1],"Jun"]=float(item[7])
            self.fc_price[item[0],item[1],"Jul"]=float(item[8])
            self.fc_price[item[0],item[1],"Aug"]=float(item[9])
            self.fc_price[item[0],item[1],"Sep"]=float(item[10])
            self.fc_price[item[0],item[1],"Oct"]=float(item[11])
            self.fc_price[item[0],item[1],"Nov"]=float(item[12])
            self.fc_price[item[0],item[1],"Dec"]=float(item[13])
        f.close()

        csvloc = os.path.join(dataloc,'Supplier Capacity.csv')
        f = open(csvloc,"r")
        myreader = csv.reader(f)
        next(myreader,None)
        self.fc_cap={}
        for item in myreader:
            self.fc_cap[item[0],item[1],"Jan"]=float(item[2])
            self.fc_cap[item[0],item[1],"Feb"]=float(item[3])
            self.fc_cap[item[0],item[1],"Mar"]=float(item[4])
            self.fc_cap[item[0],item[1],"Apr"]=float(item[5])
            self.fc_cap[item[0],item[1],"May"]=float(item[6])
            self.fc_cap[item[0],item[1],"Jun"]=float(item[7])
            self.fc_cap[item[0],item[1],"Jul"]=float(item[8])
            self.fc_cap[item[0],item[1],"Aug"]=float(item[9])
            self.fc_cap[item[0],item[1],"Sep"]=float(item[10])
            self.fc_cap[item[0],item[1],"Oct"]=float(item[11])
            self.fc_cap[item[0],item[1],"Nov"]=float(item[12])
            self.fc_cap[item[0],item[1],"Dec"]=float(item[13])
        f.close()

        print "Creating auxiliary data..."
        # Lead time indicators
        self.slow = {} # longest duration of reaching an FDP from the key location
        self.quick = {} # shortest duration of reaching an FDP from the key location
        for edp in self.EDPs:
            for fdp in self.FDPs:
                if (edp,fdp,self.commodities[0]) in self.cost.keys():
                    self.slow[edp,fdp] = self.dur[edp,fdp,self.commodities[0]]
                    self.quick[edp,fdp] = self.dur[edp,fdp,self.commodities[0]]
        for dp in self.DPs:
            for fdp in self.FDPs:
                self.slow[dp,fdp] = 0
                self.quick[dp,fdp] = 999
                for edp in self.EDPs:
                    if (edp,fdp) not in self.slow.keys():
                        continue
                    if (dp,edp,self.commodities[0]) in self.cost.keys():
                        ds = self.slow[edp,fdp] + self.dur[dp,edp,self.commodities[0]]
                        dq = self.quick[edp,fdp] + self.dur[dp,edp,self.commodities[0]]
                        if ds > self.slow[dp,fdp]:
                            self.slow[dp,fdp] = ds
                        if dq < self.quick[dp,fdp]:
                            self.quick[dp,fdp] = dq
        for arc in self.proccap.keys(): # arc= (src,ndp,com)
            for fdp in self.FDPs:
                if (arc[0],arc[1],fdp) not in self.slow.keys(): # NB: Commodity doesn't matter
                    self.slow[arc[0],arc[1],fdp] = 0
                    self.quick[arc[0],arc[1],fdp] = 999
                    if arc[1] in self.ISs:
                        for dp in self.DPs:
                            if (dp,fdp) not in self.slow.keys():
                                continue
                            if (arc[1],dp,arc[2]) in self.cost.keys():
                                ds = self.dur[arc[1],dp,arc[2]] + self.slow[dp,fdp]
                                dq = self.dur[arc[1],dp,arc[2]] + self.quick[dp,fdp]
                                if  ds > self.slow[arc[0],arc[1],fdp]:
                                    self.slow[arc[0],arc[1],fdp] = ds
                                if  dq < self.quick[arc[0],arc[1],fdp]:
                                    self.quick[arc[0],arc[1],fdp] = dq
                    if arc[1] in self.RSs:
                        for i in (self.DPs+self.EDPs):
                            if (i,fdp) not in self.slow.keys():
                                continue
                            if (arc[1],i,arc[2]) in self.cost.keys():
                                ds = self.dur[arc[1],i,arc[2]] + self.slow[i,fdp]
                                dq = self.dur[arc[1],i,arc[2]] + self.quick[i,fdp]
                                if  ds > self.slow[arc[0],arc[1],fdp]:
                                    self.slow[arc[0],arc[1],fdp] = ds
                                if  dq < self.quick[arc[0],arc[1],fdp]:
                                    self.quick[arc[0],arc[1],fdp] = dq
                    if arc[1] in self.LSs:
                        for edp in self.EDPs:
                            if (edp,fdp) not in self.slow.keys():
                                continue
                            if (arc[1],edp,arc[2]) in self.cost.keys():
                                ds = self.dur[arc[1],edp,arc[2]] + self.slow[edp,fdp]
                                dq = self.dur[arc[1],edp,arc[2]] + self.quick[edp,fdp]
                                if  ds > self.slow[arc[0],arc[1],fdp]:
                                    self.slow[arc[0],arc[1],fdp] = ds
                                if  dq < self.quick[arc[0],arc[1],fdp]:
                                    self.quick[arc[0],arc[1],fdp] = dq
                    if arc[1] in self.LMs:
                        if (arc[1],fdp,arc[2]) in self.cost.keys():
                            self.slow[arc[0],arc[1],fdp] = self.dur[arc[1],fdp,arc[2]]
                            self.quick[arc[0],arc[1],fdp] = self.dur[arc[1],fdp,arc[2]]
                    if self.quick[arc[0],arc[1],fdp] != 999: # FDP can be reached through this source
                        self.slow[arc[0],arc[1],fdp] += self.dur[arc]
                        self.quick[arc[0],arc[1],fdp] += self.dur[arc]
                    else: # If FDP can't be reached, remove the initialised values (0 for slow, 999 for quick)
                        self.slow.pop((arc[0],arc[1],fdp),None)
                        self.quick.pop((arc[0],arc[1],fdp),None)
                    # NB1: slow[src,ndp,fdp] now captures the duration of supplying the fdp through the slowest route possible
                    # NB2: quick[src,ndp,fdp] now captures the duration of supplying the fdp through the fastest route possible
        for key in self.slow.keys():
            if len(key) < 3:
                self.slow.pop(key,None)
                self.quick.pop(key,None)
                # We only need (src,ndp,fdp)-keys to define the LT measures
        for key in self.slow.keys(): # Only contains (src,ndp,fdp) keys now
            if (key[0],key[1]) in self.slow.keys():
                continue
            self.slow[key[0],key[1]] = max(self.slow[key[0],key[1],fdp] for fdp in self.FDPs if (key[0],key[1],fdp) in self.slow.keys())
            self.quick[key[0],key[1]] = max(self.quick[key[0],key[1],fdp] for fdp in self.FDPs if (key[0],key[1],fdp) in self.slow.keys())
        for key in self.slow.keys():
            if len(key) == 3:
                self.slow.pop(key,None)
                self.quick.pop(key,None)

        # Location type classification
        self.type = {}
        for i in self.ISs:
            self.type[i] = "International Supplier"
        for i in self.RSs:
            self.type[i] = "Regional Supplier"
        for i in self.LSs:
            self.type[i] = "Local Supplier"
        for i in self.LMs:
            self.type[i] = "Local Market (C&V)"
        for i in self.DPs:
            self.type[i] = "Discharge Port"
        for i in self.EDPs:
            self.type[i] = "Extended Delivery Point"
        for i in self.FDPs:
            self.type[i] = "Final Delivery Point"

        # Statistics
        self.stats = {}
        self.PC_I = {}
        self.stats["Procurement Costs [Int]"] = self.PC_I
        self.PC_L = {}
        self.stats["Procurement Costs [Loc]"] = self.PC_L
        self.PC_CV = {}
        self.stats["Procurement Costs [C&V]"] = self.PC_CV
        self.PC_R = {}
        self.stats["Procurement Costs [Reg]"] = self.PC_R
        self.PC = {}
        self.stats["Procurement Costs [Tot]"] = self.PC
        self.TR_OC = {}
        self.stats["Transportation Costs [Ocean]"] = self.TR_OC
        self.TR_OL = {}
        self.stats["Transportation Costs [Overland]"] = self.TR_OL
        self.TR_IL = {}
        self.stats["Transportation Costs [Inland]"] = self.TR_IL
        self.TR = {}
        self.stats["Transportation Costs [Total]"] = self.TR
        self.HC = {}
        self.stats["Handling Costs"] = self.HC
        self.ODOC_CV = {}
        self.stats["ODOC Costs [C&V]"] = self.ODOC_CV
        self.ODOC_F = {}
        self.stats["ODOC Costs [Food]"] = self.ODOC_F
        self.ODOC = {}
        self.stats["ODOC Costs [Total]"] = self.ODOC
        self.DOC = {}
        self.stats["Direct Operational Costs"] = self.DOC
        self.DSC = {}
        self.stats["Direct Support Costs"] = self.DSC
        self.TDC = {}
        self.stats["Total Direct Costs"] = self.TDC
        self.ISC = {}
        self.stats["Indirect Support Costs"] = self.ISC
        self.TC = {}
        self.stats["Total Costs"] = self.TC
        self.NVS = {}
        self.stats["Nutritional Value Score"] = self.NVS
        self.COMS = {}
        self.stats["Commodities (#)"] = self.COMS
        self.GROUPS = {}
        self.stats["Food Groups (#)"] = self.GROUPS
        self.FCS = {}
        self.stats["Food Consumption Score"] = self.FCS
        self.KCAL = {}
        self.stats["Energy Supplied [Total]"] = self.KCAL
        self.PROT = {}
        self.stats["Energy Supplied [Protein]"] = self.PROT
        self.FAT = {}
        self.stats["Energy Supplied [Fat]"] = self.FAT
        self.MT_I = {}
        self.stats["Procured MT [Int]"] = self.MT_I
        self.MT_L = {}
        self.stats["Procured MT [Loc]"] = self.MT_L
        self.MT_CV = {}
        self.stats["Procured MT [C&V]"] = self.MT_CV
        self.MT_R = {}
        self.stats["Procured MT [Reg]"] = self.MT_R
        self.MT = {}
        self.stats["Procured MT [Tot]"] = self.MT
        self.LTsum = {}
        self.stats["Lead Time"] =self.LTsum # will be reassigned in prep() because lead times are a bit of an exception

        # Secondary Objectives
        self.objectives=[] # These are the statistics that are interesting for trade-off graphs
        self.objectives.append("NVS (Min)")
        self.objectives.append("NVS (Avg)")
        self.objectives.append("NVS (% Supplied)")
        self.objectives.append("C&V (%)")
        self.objectives.append("Loc (%)")
        self.objectives.append("Lead Time (Avg)")
        self.objectives.append("Lead Time (Max)")
        self.objectives.append("Kcal (Avg)")
        self.base = {} # Some standard values for trade-off analyses
        self.base["NVS (Min)"] = [5.5,11,.5]
        self.base["NVS (Avg)"] = [5.5,11,.5]
        self.base["NVS (% Supplied)"] = [50,100,5]
        self.base["C&V (%)"] = [0,100,10]
        self.base["Loc (%)"] = [0,100,10]
        self.base["Lead Time (Avg)"] = [0,100,10]
        self.base["Lead Time (Max)"] = [0,200,10]
        self.base["Kcal (Avg)"] = [1500,2100,100]

        # Create indices for commodities and locations
        self.comindex = range(len(self.commodities))

        # Store variables for later use
        print "Storing data inputs..."
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'data')
        path = os.path.join(dest_dir, 'data.pickle')
        f = open(path,'wb')
        pickle.dump(vars(self),f)
        f.close()
        # Ticking the box
        path = os.path.join(dest_dir, 'check.csv')
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow([1])
        out.close()

    def load_quick(self):
        '''
        Load data from previous session.
        There's no need to reprocess the data if nothing changed
        '''

        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'data')
        path = os.path.join(dest_dir, 'data.pickle')
        f = open(path,'rb')
        M = pickle.load(f)
        for var in M.items():
            setattr(self, var[0], var[1])
        f.close()

    def prep(self):
        '''
        Setup the mathematical model by preparing all the constraints and variables that will not change between scenarios.
        Doing so saves a lot of time when analysing multiple scenarios.
        Depending on the size of the problem, this may save anything from 3 seconds to several minutes per scenario.
        '''

        print "Preparing the model:"
        self.genstart = time.time()

        # Create the subset of self.periods in which we are allowed to make decisions
        self.horizon = list(self.periods)
        for t in self.periods:
            if t != self.tstart.get():
                self.horizon.remove(t)
            else:
                break
        for t in reversed(self.periods):
            if t != self.tend.get():
                self.horizon.remove(t)
            else:
                break
        self.hor = range(len(self.horizon)) # This list contains indices to self.horizon rather than month names

        print "Creating space-time network..."
        # Create network flows for specified time period
        self.arcs = {}
        for arc in self.cost.keys():
            for t in self.hor:
                self.arcs[arc[0],arc[1],arc[2],t] = self.cost[arc]

        self.frame_right.update_idletasks()

        # Overwrite prices with forecasts if available
        if self.useforecasts.get()==1:
            # Int'l forecast
            for key in self.proccap.keys(): # key = (src, ndp, com)
                ctry = key[0][:-6]
                if key not in self.date.keys():
                    continue
                if (ctry,key[2],self.date[key]) in self.fc_price.keys():
                    for t in self.hor:
                        mon1 = self.horizon[t][0:3]
                        self.arcs[key[0],key[1],key[2],t] = self.cost[key] / self.fc_price[ctry,key[2],self.date[key]] * self.fc_price[ctry,key[2],mon1]
                        # NB: The original cost is divided by the seasonality index of the "as of" month (to get the deseasonalised base price) and then multiplied by the seasonality factor of the current month
            # CBT forecast
            for m in self.LMs:
                for k in self.commodities:
                    if (m,k,"Jan") in self.fc_price.keys():
                        for t in self.hor:
                            mon1 = self.horizon[t][0:3]
                            self.arcs['Local Markets - C&V',m,k,t] = self.fc_price[m,k,mon1]


        # Create movements across time (for each transshipment node)
        self.arr = {}
        for i in (self.ISs + self.RSs + self.LSs + self.LMs + self.DPs + self.EDPs):
            for k in self.commodities:
                for t in self.hor:
                    self.arr[i,k,t] = []  # Predefine arrivals of k in i at t
                    # NB 1: regardless of the origin of the arrival
                    # NB 2: as with inventories, predefining makes the constraints easier to define
        # Set up inventories
        for k in self.commodities:
            for i in (self.EDPs+self.DPs):
                self.dur[i,i,k] = 30
                for t in self.hor:
                    self.arcs[i,i,k,t] = 0 # EDPs and DPs can hold inventory -> storage costs are incurred for each month (covered under handling costs)
                    if (t+1) in self.hor:
                        self.arr[i,k,t+1].append([i,i,k,t])
        # Set up the dictionary with arrivals
        for key in self.cost.keys():
            if key[1] in self.FDPs:
                continue
            d = int(self.dur[key])
            temp = d % 30
            if temp <= 20: # ~~~~~~~~~~~~~ THE 20 MIGHT NEED SOME FINETUNING ~~~~~~~~~~
                T = (d-temp)/30
            else:
                T = (d-temp)/30 + 1
            if T > 0:
                for t in self.hor[:-T]:
                    self.arr[key[1],key[2],t+T].append([key[0],key[1],key[2],t])
            else:
                for t in self.hor:
                    self.arr[key[1],key[2],t].append([key[0],key[1],key[2],t])
        # NB 1 : the dictionary self.arr[i,k,t] now captures all k flowing into i at time t    (excluding the inv[i,k,t] from the data)
        # NB 2 : the corresponding outflow is captured by the [i,j,k,t] keys from self.arc


        print "Creating decision variables..."
        self.F = LpVariable.dicts('Flow',self.arcs,0,None,LpContinuous) # Note that variables are created 'sparse', i.e. only relevant arcs are included
        self.R = LpVariable.dicts('Ration',(self.commodities,self.hor),0,None,LpContinuous)
        self.CV = LpVariable.dicts("C&V",(self.FDPs,self.hor),0,None,LpContinuous) # Cash & Voucher component of the basket
        self.K = LpVariable.dicts("Commodity",(self.commodities,self.hor),0,1,LpBinary) # Auxiliary variable
        self.G = LpVariable.dicts("FoodGroup",(self.foodgroups,self.hor),0,1,LpBinary) # Auxiliary variable
        self.FCSG = LpVariable.dicts("FCSGroup",(self.fcsgroups,self.hor),0,1,LpBinary) # Auxiliary variable
        self.S = LpVariable.dicts("Shortfall",(self.nutrients,self.hor),0,1,LpContinuous)
        self.O = LpVariable.dicts("Overshoot",(self.nutrients,self.hor),0,None,LpContinuous) # Auxiliary variable
        self.SFI = LpVariable.dicts("ShortfallIndicator",(self.nutrients,self.hor),0,1,LpBinary) # Auxiliary variable
        self.P = LpVariable.dicts("Procured",(self.slow,self.hor),0,1,LpBinary) # Auxiliary variable
        self.LT = LpVariable.dicts("LeadTime",self.hor,0,None,LpContinuous) # Auxiliary variable
        self.LTmax = LpVariable("LTmax",0,None,LpContinuous)
        self.stats["Lead Time"] =self.LT # initialised as LTsum, but LT makes more sense to track

        print "Creating statistics..."
        # Procurement Costs
        for t in self.hor:
            self.PC_I[t] = lpSum([self.F[proc[0]]*proc[1] for proc in self.arcs.items() if proc[0][0] in self.sources and proc[0][1] in self.ISs and proc[0][3]==t]) # Procurement costs (int)
            self.PC_L[t] = lpSum([self.mod_loc * self.F[proc[0]]*proc[1] for proc in self.arcs.items() if proc[0][0] in self.sources and proc[0][1] in self.LSs and proc[0][3]==t]) # Procurement costs (loc)
            self.PC_CV[t] = lpSum([self.mod_cbt * self.F[proc[0]]*proc[1] for proc in self.arcs.items() if proc[0][0] in self.sources and proc[0][1] in self.LMs and proc[0][3]==t]) # Procurement costs (C&V)
            self.PC_R[t] = lpSum([self.mod_reg * self.F[proc[0]]*proc[1] for proc in self.arcs.items() if proc[0][0] in self.sources and proc[0][1] in self.RSs and proc[0][3]==t]) # Procurement costs (reg)
            self.PC[t] = self.PC_I[t] + self.PC_L[t] + self.PC_CV[t] + self.PC_R[t] # Total procurement costs
            # NB: We look at self.arcs.items() because procurement prices may differ between periods (forecast)
        self.PC_I["Total"] = sum(self.PC_I[t] for t in self.hor)
        self.PC_I["Average"] = self.PC_I["Total"]/float(len(self.hor))
        self.PC_L["Total"] = sum(self.PC_L[t] for t in self.hor)
        self.PC_L["Average"] = self.PC_L["Total"]/float(len(self.hor))
        self.PC_CV["Total"] = sum(self.PC_CV[t] for t in self.hor)
        self.PC_CV["Average"] = self.PC_CV["Total"]/float(len(self.hor))
        self.PC_R["Total"] = sum(self.PC_R[t] for t in self.hor)
        self.PC_R["Average"] = self.PC_R["Total"]/float(len(self.hor))
        self.PC["Total"] = sum(self.PC[t] for t in self.hor)
        self.PC["Average"] = self.PC["Total"]/float(len(self.hor))
        # TRansportation costs
        for t in self.hor:
            self.TR_OC[t] = lpSum([self.F[arc[0],arc[1],arc[2],t]*self.cost[arc] for arc in self.cost.keys() if arc[0] not in self.sources and arc[1] in self.DPs]) # OCean costs
            self.TR_OL[t] = lpSum([self.F[arc[0],arc[1],arc[2],t]*self.cost[arc] for arc in self.cost.keys() if arc[0] in self.RSs and arc[1] in self.EDPs]) # OverLand costs
            self.TR_IL[t] = lpSum([self.F[arc[0],arc[1],arc[2],t]*self.cost[arc] for arc in self.cost.keys() if ((arc[0] not in self.RSs and arc[1] in self.EDPs) or arc[1] in self.FDPs)]) # InLand costs
            self.TR[t] = self.TR_OC[t] + self.TR_OL[t] + self.TR_IL[t] # Total transportation costs
            # NB: We look at self.cost.keys() rather than self.arc.items() because the transportation cost doesn't change over time - this saves a lot of computation time
        self.TR_OC["Total"] = sum(self.TR_OC[t] for t in self.hor)
        self.TR_OC["Average"] = self.TR_OC["Total"]/float(len(self.hor))
        self.TR_OL["Total"] = sum(self.TR_OL[t] for t in self.hor)
        self.TR_OL["Average"] = self.TR_OL["Total"]/float(len(self.hor))
        self.TR_IL["Total"] = sum(self.TR_IL[t] for t in self.hor)
        self.TR_IL["Average"] = self.TR_IL["Total"]/float(len(self.hor))
        self.TR["Total"] = sum(self.TR[t] for t in self.hor)
        self.TR["Average"] = self.TR["Total"]/float(len(self.hor))
        # Handling Costs
        self.LOAD = {}
        self.LOAD_F = {}
        self.LOAD_CV = {}
        for t in self.hor:
            for i in (self.DPs+self.EDPs):
                self.LOAD[i,t] = lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==i) + lpSum(self.F[i,i,k,t] for k in self.commodities) # Flow originating from transshipment point i
                self.HC[i,t] = lpSum(self.F[arc[0],arc[1],arc[2],t]*self.hc[i] for arc in self.cost.keys() if arc[0]==i) + lpSum(self.F[i,i,k,t]*self.sc[i] for k in self.commodities)
                # NB: capturing flows arriving in an (E)DP is messy due to lead times, but we know that outflow[t] = inflow[t]
            for i in (self.FDPs):
                self.LOAD_F[i,t] = lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[1]==i and arc[0] not in self.LMs) # Flow (non-C&V) arriving in fdp i
                self.LOAD_CV[i,t] = lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[1]==i and arc[0] in self.LMs) # Flow (C&V) arriving in fdp i
                self.LOAD[i,t] = self.LOAD_F[i,t] + self.LOAD_CV[i,t]
                self.HC[i,t] = self.hc[i]*self.LOAD_F[i,t] # Distribution costs are not incurred for C&V
            self.HC[t] = sum(self.HC[i,t] for i in (self.DPs+self.EDPs+self.FDPs))
        self.HC["Total"] = sum(self.HC[t] for t in self.hor)
        self.HC["Average"] = self.HC["Total"]/float(len(self.hor))
        # ODOC costs
        for t in self.hor:
            self.ODOC_F[t] = lpSum([self.F[arc[0],arc[1],arc[2],t]*self.odocF for arc in self.cost.keys() if arc[0] in self.sources and arc[1] not in self.LMs]) # ODOC costs (food)
            self.ODOC_CV[t] = self.PC_CV[t] * self.odocCV # ODOC costs (C&V)
            self.ODOC[t] = self.ODOC_F[t] + self.ODOC_CV[t]
        self.ODOC_CV["Total"] = sum(self.ODOC_CV[t] for t in self.hor)
        self.ODOC_CV["Average"] = self.ODOC_CV["Total"]/float(len(self.hor))
        self.ODOC_F["Total"] = sum(self.ODOC_F[t] for t in self.hor)
        self.ODOC_F["Average"] = self.ODOC_F["Total"]/float(len(self.hor))
        self.ODOC["Total"] = sum(self.ODOC[t] for t in self.hor)
        self.ODOC["Average"] = self.ODOC["Total"]/float(len(self.hor))
        # Total Costs
        for t in self.hor:
            self.DOC[t] = self.PC[t] + self.TR[t] + self.HC[t] + self.ODOC[t] # Direct Operational Costs
            self.DSC[t] = self.DOC[t]*self.dsc # Direct Support Costs
            self.TDC[t] = self.DOC[t] + self.DSC[t] # Total Direct Costs
            self.ISC[t] = self.TDC[t]*self.isc # Indirect Support Costs
            self.TC[t] = self.TDC[t] + self.ISC[t] # Total Costs of the operation
        self.DOC["Total"] = sum(self.DOC[t] for t in self.hor)
        self.DOC["Average"] = self.DOC["Total"]/float(len(self.hor))
        self.DSC["Total"] = sum(self.DSC[t] for t in self.hor)
        self.DSC["Average"] = self.DSC["Total"]/float(len(self.hor))
        self.TDC["Total"] = sum(self.TDC[t] for t in self.hor)
        self.TDC["Average"] = self.TDC["Total"]/float(len(self.hor))
        self.ISC["Total"] = sum(self.ISC[t] for t in self.hor)
        self.ISC["Average"] = self.ISC["Total"]/float(len(self.hor))
        self.TC["Total"] = sum(self.TC[t] for t in self.hor)
        self.TC["Average"] = self.TC["Total"]/float(len(self.hor))
        # Other statistics
        for t in self.hor:
            self.NVS[t] = len(self.nutrients)+lpSum(self.S[l][t]*-1 for l in self.nutrients) # Nutritional Value Score
            self.COMS[t] = lpSum(self.K[k][t] for k in self.commodities) # Amount of unique commodities
            self.GROUPS[t] = lpSum(self.G[g][t] for g in self.foodgroups) # Amount of unique food groups
            self.FCS[t] = lpSum(7*self.weight[g]*self.FCSG[g][t] for g in self.fcsgroups) # Food Contribution Score
            self.KCAL[t] = lpSum([float(self.nutval[k,"ENERGY (kcal)"])/100*self.R[k][t] for k in self.commodities]) # Amount of kcal supplied per beneficiary
            self.PROT[t] = lpSum([4.1*float(self.nutval[k,"PROTEIN (g)"])/100*self.R[k][t] for k in self.commodities]) # Amount of kcal supplied through proteins
            self.FAT[t] = lpSum([8.8*float(self.nutval[k,"FAT    (g)"])/100*self.R[k][t] for k in self.commodities]) # Amount of kcal supplied through fats
            self.MT_I[t] = lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.ISs) # Total amount of mt purchased (int)
            self.MT_R[t] = lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.RSs) # Total amount of mt purchased (reg)
            self.MT_L[t] = lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LSs) # Total amount of mt purchased (loc)
            self.MT_CV[t] = lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LMs) # Total amount of mt purchased (c&v)
            self.MT[t] = self.MT_I[t] + self.MT_R[t] + self.MT_L[t] + self.MT_CV[t] # Total amount of mt purchased
            self.LTsum[t] = lpSum(self.F[arc[0],arc[1],arc[2],t]*self.dur[arc] for arc in self.cost.keys()) # Sum of lead times    NB: Has to be divided by MT[t] to get the average lead time
        # Could add weighted functions, minimum NVS, avg NVS, discounts, etc.??
        self.NVS["Total"] = sum(self.NVS[t] for t in self.hor)
        self.COMS["Total"] = sum(self.COMS[t] for t in self.hor)
        self.GROUPS["Total"] = sum(self.GROUPS[t] for t in self.hor)
        self.FCS["Total"] = sum(self.FCS[t] for t in self.hor)
        self.KCAL["Total"] = sum(self.KCAL[t] for t in self.hor)
        self.PROT["Total"] = sum(self.PROT[t] for t in self.hor)
        self.FAT["Total"] = sum(self.FAT[t] for t in self.hor)
        self.MT_I["Total"] = sum(self.MT_I[t] for t in self.hor)
        self.MT_I["Average"] = self.MT_I["Total"]/float(len(self.hor))
        self.MT_R["Total"] = sum(self.MT_R[t] for t in self.hor)
        self.MT_R["Average"] = self.MT_R["Total"]/float(len(self.hor))
        self.MT_L["Total"] = sum(self.MT_L[t] for t in self.hor)
        self.MT_L["Average"] = self.MT_L["Total"]/float(len(self.hor))
        self.MT_CV["Total"] = sum(self.MT_CV[t] for t in self.hor)
        self.MT_CV["Average"] = self.MT_CV["Total"]/float(len(self.hor))
        self.MT["Total"] = sum(self.MT[t] for t in self.hor)
        self.MT["Average"] = self.MT["Total"]/float(len(self.hor))
        self.LTsum["Total"] = sum(self.LTsum[t] for t in self.hor)
        self.LTsum["Average"] = self.LTsum["Total"]/float(len(self.hor))

        print "Creating model constraints..."
        self.CORE = {} # Used to store general constraints
        count = 0

        # Lead time tracking
        for key in self.proccap.keys(): # key = (src, ndp, com)
            for t in self.hor:
                self.CORE[count]= self.F[key[0],key[1],key[2],t] <= 1000000000 * self.P[key[0],key[1]][t]
                self.CORE[count+1]= self.LT[t] >= self.P[key[0],key[1]][t] * self.quick[key[0],key[1]]
                count+=2
        for t in self.hor:
            self.CORE[count]= self.LTmax >= self.LT[t]
            count+=1

        # Shortfall tracking
        for t in self.hor:
            for l in self.nutrients:
                self.CORE[count] = self.S[l][t] <= self.SFI[l][t]
                self.CORE[count+1] = self.O[l][t] <= (1-self.SFI[l][t])*100
                count+=2
                # NB: SFI is 1 if there's a shortfall for l at t, so now S and O can't be >0 at the same time

        # Set up Z_kt variable
        for t in self.hor:
            for k in self.commodities:
                if k == "CASH":
                    self.CORE[count]=   (self.R[k][t] >= self.K[k][t] * .01)
                else:
                    self.CORE[count]=   (self.R[k][t] >= self.K[k][t] * 1)
                self.CORE[count+1]= (self.R[k][t] <= self.K[k][t] * 10000)  # gr/ration for any commodity is =<1000 gr and >=1 gram (if included in basket) (exception for CASH commodity)
                count+=2

        # Set up G_g variable
        for t in self.hor:
            for g in self.foodgroups:
                self.CORE[count]= lpSum([self.K[k][t] for k in self.commodities if self.group[k]==g]) <= 10000*self.G[g][t]
                self.CORE[count+1]= lpSum([self.K[k][t] for k in self.commodities if self.group[k]==g]) >= self.G[g][t]
                count+=2

        # Set up FCS_g variable
        for t in self.hor:
            for g in self.fcsgroups:
                self.CORE[count]= lpSum([self.K[k][t] for k in self.commodities if self.fcs[k]==g]) <= 10000*self.FCSG[g][t]
                self.CORE[count+1]= lpSum([self.K[k][t] for k in self.commodities if self.fcs[k]==g]) >= self.FCSG[g][t]
                count+=2

        # Network Flow Constraints
        # Source nodes (bound outflow and set up lead time)
        for key in self.proccap.keys(): # key = (src, ndp, com)
            for t in self.hor:
                self.CORE[count]= self.F[key[0],key[1],key[2],t] <= self.proccap[key]
                count+=1
        for key in self.fc_cap.keys(): # key = (origin country, com, month)
            for t in self.hor:
                if self.horizon[t].startswith(key[2]):
                    self.CORE[count]= lpSum([self.F[proc[0],proc[1],proc[2],t] for proc in self.cost.keys() if proc[0].startswith(key[0]) and proc[2]==key[1] ]) <= self.fc_cap[key]
                    count+=1

        # Transhipment nodes (flow out = flow in)
        for key in self.arr.keys(): # key = (i,k,t)  where i: transshipment node. self.arr[key] is then the set of [i,j,k,t*]'s that arrive in [i,k,t]
            if key[0] in (self.DPs+self.EDPs): # includes inventory
                self.CORE[count]= lpSum(self.F[arc[0],arc[1],arc[2],arc[3]] for arc in self.arr[key]) + self.inv[key[0],key[1],self.horizon[key[2]]] == lpSum(self.F[arc[0],arc[1],arc[2],key[2]] for arc in self.cost.keys() if arc[0]==key[0] and arc[2]==key[1]) + self.F[key[0],key[0],key[1],key[2]]
            else: # doesn't include inventory
                self.CORE[count]= lpSum(self.F[arc[0],arc[1],arc[2],arc[3]] for arc in self.arr[key]) + self.inv[key[0],key[1],self.horizon[key[2]]] == lpSum(self.F[arc[0],arc[1],arc[2],key[2]] for arc in self.cost.keys() if arc[0]==key[0] and arc[2]==key[1])
            # NB: arrivals of k in i at t + whatever arrives for free == departures of k from i at t (including inventory)
            count+=1

		# Arc capacities
        for arc in self.arccap.keys(): # arc = (orig, dest, period)
            if arc[2] in self.horizon:
                t = self.horizon.index(arc[2])
                self.CORE[count]= lpSum(self.F[arc[0],arc[1],k,t] for k in self.commodities) <= self.arccap[arc[0],arc[1],arc[2]]
                count+=1

		# Node capacities
        for t in self.hor:
            for i in (self.DPs+self.EDPs):
                self.CORE[count]= lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==i) + lpSum(self.F[i,i,k,t] for k in self.commodities) <= self.nodecap[i,self.horizon[t]]
                count+=1

        self.gendur = time.time()-self.genstart
        self.n_vars = len(self.F.keys() + self.R.keys() + self.CV.keys() + self.K.keys() + self.G.keys() + self.FCSG.keys() + self.S.keys() + self.O.keys() + self.SFI.keys() + self.P.keys() + self.LT.keys()) + 1
        print "Finished preparing the model:"
        print str(self.n_vars)+" Variables & "+str(count)+" Constraints in "+str(self.gendur)+" seconds"
        print " "
        self.prepped = [self.tstart.get(),self.tend.get(),self.useforecasts.get()] # Tracks whether the data has been set up yet

    def calculate(self, NAME):
        '''
        Add user inputs to the core model (from prep(self)) and set up the optimisation model.
        The resulting LP is then sent to the solver.
        '''
        if self.prepped != [self.tstart.get(),self.tend.get(),self.useforecasts.get()]: # the general constraints have not been set yet for the current time horizon
            self.prep()
        self.calcstart = time.time()
        print "Calculating scenario: " + NAME
        print "Including the general constraints"
        self.errors = 0 # will keep track of raised errors
        self.frame_right.update_idletasks()



        # Creates the 'prob' variable to contain the problem data
        prob = LpProblem("UNWFP",LpMinimize)
        prob += self.TC["Total"]

        for key in self.CORE.keys():
            prob += self.CORE[key] # Adds the general constraints that were prepared in self.prep()
        self.n_constr = len(self.CORE.keys())
        print "Setting up constraints from user input"

        # Some months have no demand, resulting in some exceptions
        self.empty = []
        for t in self.hor:
            if sum(self.dem[self.ben.get(),i,self.horizon[t]] for i in self.FDPs) == 0 : # 'empty' month
                self.empty.append(t)
        # No basket for 'empty' months in order to better track nutritional objectives
        for t in self.empty:
            prob += lpSum(self.R[k][t] for k in self.commodities) == 0
            self.n_constr += 1
        # Nutritional objectives are redefined slightly because empty months have no food baskets
        self.KCAL["Average"] = self.KCAL["Total"]/(float(len(self.hor))-len(self.empty)) # no kcal in months without demand
        self.PROT["Average"] = self.PROT["Total"]/(float(len(self.hor))-len(self.empty)) # no protein in months without demand
        self.FAT["Average"] = self.FAT["Total"]/(float(len(self.hor))-len(self.empty)) # no fat in months without demand
        self.FCS["Average"] = self.FCS["Total"]/(float(len(self.hor))-len(self.empty)) # no FCS in months without demand
        self.GROUPS["Average"] = self.GROUPS["Total"]/(float(len(self.hor))-len(self.empty)) # no food groups in months without demand
        self.NVS["Average"] = self.NVS["Total"]/(float(len(self.hor))-len(self.empty)) # no NVS in months without demand
        self.COMS["Average"] = self.COMS["Total"]/(float(len(self.hor))-len(self.empty)) # no commodities in months without demand

        # Demand nodes (bound inflow)
        feed_days = max(self.feedingdays[self.ben.get(),k] for k in self.commodities)
        if feed_days == 0:
            feed_days = 30
            # NB: If no food basket was pre-defined for this activity, assume the default of 30 feeding days
        for i in self.FDPs:
            for k in self.commodities:
                for t in self.hor:
                    extra_demand = sum(self.dem[b,i,self.horizon[t]]*self.baskets[b,k]*self.feedingdays[b,k] for b in self.activities)
                    if self.supply_tact.get() == 1:
                        if (i,k,self.horizon[t]) in self.tact_demand.keys():
                            if self.tactboxes[i].get()==1 and self.tactboxes[k].get()==1 and self.tactboxes[self.horizon[t]].get()==1:
                                extra_demand += self.tact_demand[i,k,self.horizon[t]] * 1000000
                    prob += lpSum(1000000*self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[1]==i and arc[2]==k) >= ( self.dem[self.ben.get(),i,self.horizon[t]] * feed_days * self.R[k][t] + extra_demand ) * self.scaleup
                    #             1000000 grams/mt * mt supplied = grams supplied /month for (i,k,t)                             == amount of beneficiaries * 30 days/month * gr/day + demand from other activities
                    self.n_constr += 1

        # Variable food basket
        if self.varbasket.get()=="Fix All":
            temp = [t for t in self.hor if t not in self.empty]
            for t1 in temp:
                for t2 in temp:
                    if t1!=t2:
                        # t1 and t2 are now two non-empty time periods in the horizon
                        for k in self.commodities:
                            prob += self.R[k][t1] == self.R[k][t2]
                            # NB: Traditionally Y[k][t]=Y[k][t+1] would be used, but that would conflict with time periods where we have no demand (and thus no food basket)
                            self.n_constr += 1
        elif self.varbasket.get()=="Fix Commodities":
            temp = [t for t in self.hor if t not in self.empty]
            for t1 in temp:
                for t2 in temp:
                    if t1!=t2:
                        # t1 and t2 are now two non-empty time periods in the horizon
                        for k in self.commodities:
                            prob += self.K[k][t1] == self.K[k][t2]
                            self.n_constr += 1

        # Sensible Food Basket Constraints
        if(self.sensible.get() == 1):
            '''
            Sensible food baskets have a staple (300-500 gr of cereal), enrich this with pulses (30-130 gr),
            and add some oil for cooking (15-35 gr). Luxury goods (meat, fish, dairy) are not very cost-effective,
            and should be discouraged. Blended Foods are not meant for general food distribution, and should also be limited.
            '''
            for t in self.hor:
                if t in self.empty:
                    continue # We don't have a food basket in empty months
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "CEREALS & GRAINS") >= 250
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "CEREALS & GRAINS") <= 500
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "PULSES & VEGETABLES") >= 30
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "PULSES & VEGETABLES") <= 130
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "OILS & FATS") >= 15
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "OILS & FATS") <= 40
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "MIXED & BLENDED FOODS") <= 60
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "DAIRY PRODUCTS") <= 40
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "MEAT") <= 40
                prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == "FISH") <= 40
                prob += self.R["IODISED SALT"][t] >= 5 # To satisfy Iodine (tracked) and Sodium (not tracked) requirements
                prob += self.PROT[t] >= 0.1 * self.KCAL[t]
                prob += self.PROT[t] <= 0.2 * self.KCAL[t]
                prob += self.FAT[t] >= 0.17 * self.KCAL[t]
                self.n_constr += 14
##                prob += self.FAT[t] <= 0.50 * self.KCAL[t]
##                prob += self.S["ENERGY (kcal)"][t] <= .1
##                prob += self.S["FAT    (g)"][t] <= .1
##                prob += self.S["PROTEIN (g)"][t] <= .1

        # US IK donation
        if len(self.user_add_ik.keys()) > 0:
            self.PC_IK = {}
            self.TR_IK = {}
            self.MT_IK = {}
            self.TC_IK = {}
            for key in self.dur.keys():
                if key[0].startswith("USA"):
                    d = int(self.dur[key])
                    temp = d % 30
                    if temp <= 20:
                        T = (d-temp)/30
                    else:
                        T = (d-temp)/30 + 1
                    # NB: T now captures the time between procurement and shipping
                    break
            for t in self.hor:
                self.PC_IK[t] = lpSum([self.F[proc[0]]*(proc[1]+self.odocF) for proc in self.arcs.items() if proc[0][0].startswith("USA") and proc[0][3]==t]) # Procurement costs (IK) + ODOC costs
                self.MT_IK[t] = lpSum([self.F[key] for key in self.arcs.keys() if key[0].startswith("USA") and key[3]==t]) # Procured Metric Tonnes
                if t in self.hor[:-T]:
                    self.TR_IK[t] = self.MT_IK[t]*self.ltsh + lpSum([self.F[key[0]]*key[1] for key in self.arcs.items() if key[0][0].endswith("(USA)") and key[0][3]==t+T]) # Shipping costs (movements from LPs in USA) + LTSH costs
                else:
                    self.TR_IK[t] = 0
                self.TC_IK[t] = (self.PC_IK[t] + self.TR_IK[t]) * (1 + self.dsc) * (1 + self.isc)
            self.PC_IK["Total"] = sum(self.PC_IK[t] for t in self.hor)
            self.TR_IK["Total"] = sum(self.TR_IK[t] for t in self.hor)
            self.MT_IK["Total"] = sum(self.MT_IK[t] for t in self.hor)
            self.TC_IK["Total"] = sum(self.TC_IK[t] for t in self.hor)
        for i in self.user_add_ik.items(): # (met,mea,t) (val)
            try:
                met,mea,t,val = i[0][0],i[0][1],self.horizon.index(i[0][2]),float(i[1])
                if met == "USD":
                    if mea == "Value":
                        prob += self.TC_IK[t] >= val
                    else: # mea = Percentage
                        prob += self.TC_IK[t] >= val/100.0 * self.TC["Total"]
                else: # met = MT
                    if mea == "Value":
                        prob += self.MT_IK[t] >= val
                    else: # mea = Percentage
                        prob += self.MT_IK[t] >= val/100.0 * self.MT["Total"]
                self.n_constr += 1
            except:
                print "<<ERROR>> Could not add In-Kind Donation (metric, measure, period) (value):"
                print i
                self.errors += 1

        # Transfer Modality Constraints
        # CO-wide requirements
        for t in self.hor:
            prob += self.MT_CV["Total"] >= float(self.user_cv_min.get())/100.0 * self.MT["Total"]
            prob += self.MT_CV["Total"] <= float(self.user_cv_max.get())/100.0 * self.MT["Total"]
            self.n_constr += 2

        # Expenditure patterns
        for t in self.hor:
            for i in self.FDPs:
                val_g = lpSum(self.F[key]*self.arcs["Local Markets - C&V",key[0],key[2],key[3]] for key in self.arcs.keys() if key[0] in self.LMs and key[1]==i and key[3]==t and self.group[key[2]]=="CEREALS & GRAINS")
                val_v = lpSum(self.F[key]*self.arcs["Local Markets - C&V",key[0],key[2],key[3]] for key in self.arcs.keys() if key[0] in self.LMs and key[1]==i and key[3]==t and self.group[key[2]]=="PULSES & VEGETABLES")
                val_c = lpSum(self.F[key]*self.arcs["Local Markets - C&V",key[0],key[2],key[3]] for key in self.arcs.keys() if key[0] in self.LMs and key[1]==i and key[3]==t and key[2]=="CASH")
                val_o = lpSum(self.F[key]*self.arcs["Local Markets - C&V",key[0],key[2],key[3]] for key in self.arcs.keys() if key[0] in self.LMs and key[1]==i and key[3]==t and self.group[key[2]]!="CEREALS & GRAINS" and self.group[key[2]]!="PULSES & VEGETABLES" and key[2]!="CASH")
                prob += self.CV[i][t] == val_g + val_v + val_c + val_o
                self.n_constr += 1
                if self.modality.get() == "Cash":
                    prob += val_g >= float(self.exp_pattern["Cereals and Grains",0].get())/100 * self.CV[i][t]
                    prob += val_g <= float(self.exp_pattern["Cereals and Grains",1].get())/100 * self.CV[i][t]
                    prob += val_v >= float(self.exp_pattern["Vegetables and Fruits",0].get())/100 * self.CV[i][t]
                    prob += val_v <= float(self.exp_pattern["Vegetables and Fruits",1].get())/100 * self.CV[i][t]
                    prob += val_c >= float(self.exp_pattern["Non-Food Items",0].get())/100 * self.CV[i][t]
                    prob += val_c <= float(self.exp_pattern["Non-Food Items",1].get())/100 * self.CV[i][t]
                    prob += val_o >= float(self.exp_pattern["Other Food Items",0].get())/100 * self.CV[i][t]
                    prob += val_o <= float(self.exp_pattern["Other Food Items",1].get())/100 * self.CV[i][t]
                    self.n_constr += 8

        # FDP-specific requirements
        for i in self.user_modality.items(): # i = (fdp,t) (min,max)
            try:
                f,t,mn,mx = i[0][0],self.horizon.index(i[0][1]),float(i[1][0]),float(i[1][1])
                if f != "All":
                    prob += self.LOAD_CV[f,t] <= mx/100.0 * self.LOAD[f,t]
                    prob += self.LOAD_CV[f,t] >= mn/100.0 * self.LOAD[f,t]
                    self.n_constr += 2
                else:
                    for i in self.FDPs:
                        prob += self.LOAD_CV[i,t] <= mx/100.0 * self.LOAD[i,t]
                        prob += self.LOAD_CV[i,t] >= mn/100.0 * self.LOAD[i,t]
                        self.n_constr += 2
            except:
                print "<<ERROR>> Could not add Transfer Modality Decision (fdp, t, min, max):"
                print i
                self.errors += 1

        # Included C&V purchases
        for i in self.user_add_cv.items(): # i = (LM, com, t) (mt)
            try:
                t = self.horizon.index(i[0][2])
                prob += self.F["Local Markets - C&V", i[0][0], i[0][1], t] >= float(i[1])
                self.n_constr += 1
            except:
                print "<<ERROR>> Could not add Local Procurement decision (LM, com, t, mt):"
                print i
                self.errors += 1

        # Excluded C&V purchases
        for i in self.user_ex_cv.items(): # i = (LM, com) ([t])
            for month in i[1]:
                try:
                    t = self.horizon.index(month)
                    if i[0][0]=="Any":
                        if i[0][1]=="Any":
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LMs]) == 0
                        else:
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LMs and arc[2]==i[0][1]]) == 0
                    else:
                        if i[0][1]=="Any":
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LMs and arc[1]==i[0][0]]) == 0
                        else:
                            prob += self.F["Local Markets - C&V", i[0][0], i[0][1], t] == 0
                    self.n_constr += 1
                except:
                    print "<<ERROR>> Could not exclude Local Procurement decision (LM/LS, com, t):"
                    print i
                    print "      for t= " + month
                    self.errors += 1

        # Total mt restriction for automated analysis: Adjust Transfer Modality
        if self.totalmt != "":
            prob += self.MT["Total"] <=  self.totalmt*1.01
            prob += self.MT["Total"] >=  self.totalmt*0.99
            self.n_constr += 2
            # NB: In order to adjust the C&V% the model would buy ridiculous amounts of C&V products
            #     this constraint prevents the model from purchasing superfluous commodities

        # Procurement allocation constraints
        for t in self.hor:
            prob += self.MT_I["Total"] >= float(self.user_int_min.get())/100.0 * self.MT["Total"]
            prob += self.MT_I["Total"] <= float(self.user_int_max.get())/100.0 * self.MT["Total"]
            prob += self.MT_R["Total"] >= float(self.user_reg_min.get())/100.0 * self.MT["Total"]
            prob += self.MT_R["Total"] <= float(self.user_reg_max.get())/100.0 * self.MT["Total"]
            prob += self.MT_L["Total"] >= float(self.user_loc_min.get())/100.0 * self.MT["Total"]
            prob += self.MT_L["Total"] <= float(self.user_loc_max.get())/100.0 * self.MT["Total"]
            self.n_constr += 6

        # Supply (all) demand for selected beneficiary type
        for t in self.hor:
            for l in self.nutrients:
                prob += lpSum([self.nutval[k,l]/100*self.R[k][t] for k in self.commodities]) == self.nutreq[self.ben.get(),l] - self.S[l][t]*self.nutreq[self.ben.get(),l] + self.O[l][t]*self.nutreq[self.ben.get(),l]
                #                      nutrient/gr * gr/ration = nutrient/ration supplied         ==   nutrient/ration requirement - shortfalls + overshoot (slack variables)
                self.n_constr += 1

        # Input Commodity Constraints
        for item in self.user_add_com.items(): # k = (com) (minrat, maxrat)
            k = [item[0], item[1][0], item[1][1]] # easier notation
            for t in self.hor:
                if t in self.empty:
                    continue
                if k[0] in self.supcom: # General Commodity added
                    if k[1] != "N/A":
                        prob += lpSum(self.R[c][t] for c in self.commodities if self.sup[c]==k[0]) >= float(k[1])
                        self.n_constr += 1
                    if k[2] != "N/A":
                        prob += lpSum(self.R[c][t] for c in self.commodities if self.sup[c]==k[0]) <= float(k[2])
                        self.n_constr += 1
                else: # Specific Commodity added
                    if k[2] != "N/A":
                        prob += self.R[k[0]][t] <= float(k[2])
                        self.n_constr += 1
                    if k[1] != "N/A" :
                        prob += self.R[k[0]][t] >= float(k[1])
                        self.n_constr += 1

        # Fix Food Basket Constraints
        for k in self.food2fix: # k = (com, gr/ration)
            for t in self.hor:
                if t in self.empty:
                    continue
                # NB: self.ration, self.remove, and self.replace are used for automatic scenario analyses
                if self.ration == "":
                    if k[0] == self.remove:
                        continue
                    if len(self.replace)==0:
                        prob += self.R[k[0]][t] == float(k[1]) # This is what happens if this scenario is not part of an automatic analysis
                    elif k[0]==self.replace[0]:
                        prob += self.R[self.replace[1]][t] == float(k[1])
                    else:
                        prob += self.R[k[0]][t] == float(k[1])
                    self.n_constr += 1
                else:
                    prob += self.R[k[0]][t] >= float(k[1])*(1-self.ration)
                    prob += self.R[k[0]][t] <= float(k[1])*(1+self.ration)
                    self.n_constr += 2

        # Exclude Commodity
        for com in self.user_ex_com:
            if com in self.supcom:
                for k in self.commodities:
                    if self.sup[k] == com:
                        for t in self.hor:
                            prob += self.K[k][t] == 0
                            self.n_constr += 1
            else:
                for t in self.hor:
                    prob += self.K[com][t] == 0
                    self.n_constr += 1

        # Include Procurement (int)
        for item in self.user_add_proc_int.items(): # item = (country, inco, ndp, com) (mt,t)
            c,i,l,k,q = item[0][0], item[0][1], item[0][2], item[0][3], float(item[1][0]) # easier notation
            for month in item[1][1]:
                try: # It's tricky to guide the input in such a way that the result is always a valid procurement decision, hence the try/except
                    t = self.horizon.index(month)
                    if c=="Any":
                        if i=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in (self.ISs+self.RSs) and arc[2]==k]) >= q
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1]==l and arc[2]==k]) >= q
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].endswith(i) and arc[1] in (self.ISs+self.RSs) and arc[2]==k]) >= q
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].endswith(i) and arc[1]==l and arc[2]==k]) >= q
                    else:
                        if i=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1] in (self.ISs+self.RSs) and arc[2]==k]) >= q
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1]==l and arc[2]==k]) >= q
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0]==c+" - "+i and arc[1] in (self.ISs+self.RSs) and arc[2]==k]) >= q
                            else:
                                prob += self.F[c+" - "+i,l,k,t] >= q
                    self.n_constr += 1
                except:
                    print "<<ERROR>> Could not add Procurement Decision (int):"
                    print item
                    print "    for t= " + month
                    self.errors += 1

        # Include Procurement (loc)
        for item in self.user_add_proc_loc.items(): # item = (country, inco, ndp, com) (mt,t)
            c,i,l,k,q = item[0][0], item[0][1], item[0][2], item[0][3], float(item[1][0]) # easier notation
            for month in item[1][1]:
                try: # It's tricky to guide the input in such a way that the result is always a valid procurement decision, hence the try/except
                    t = self.horizon.index(month)
                    if c=="Any":
                        if i=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LSs and arc[2]==k]) >= q
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1]==l and arc[2]==k]) >= q
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].endswith(i) and arc[1] in self.LSs and arc[2]==k]) >= q
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].endswith(i) and arc[1]==l and arc[2]==k]) >= q
                    else:
                        if i=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1] in self.LSs and arc[2]==k]) >= q
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1]==l and arc[2]==k]) >= q
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0]==c+" - "+i and arc[1] in self.LSs and arc[2]==k]) >= q
                            else:
                                prob += self.F[c+" - "+i,l,k,t] >= q
                    self.n_constr += 1
                except:
                    print "<<ERROR>> Could not add Procurement Decision (loc):"
                    print item
                    print "    for t= " + month
                    self.errors += 1

        # Exclude Source (int)
        for item in self.user_ex_proc_int.items(): # src = (country, ndp, com) (t)
            c,l,k = item[0][0], item[0][1], item[0][2] # easier notation
            for month in item[1]:
                try:
                    t = self.horizon.index(month)
                    if c=="Any":
                        if k=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in (self.ISs+self.RSs)]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1]==l]) == 0
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in (self.ISs+self.RSs) and arc[2]==k]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1]==l and arc[2]==k]) == 0
                    else:
                        if k=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1] in (self.ISs+self.RSs)]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1]==l]) == 0
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1] in (self.ISs+self.RSs) and arc[2]==k]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1]==l and arc[2]==k]) == 0
                    self.n_constr += 1
                except:
                    print "<<ERROR>> Could not exclude Procurement Decision:"
                    print item
                    print "    for t= " + month
                    self.errors += 1

        # Exclude Source (loc)
        for item in self.user_ex_proc_loc.items(): # src = (country, ndp, com) (t)
            c,l,k = item[0][0], item[0][1], item[0][2] # easier notation
            for month in item[1]:
                try:
                    t = self.horizon.index(month)
                    if c=="Any":
                        if k=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LSs]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1]==l]) == 0
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1] in self.LSs and arc[2]==k]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[1]==l and arc[2]==k]) == 0
                    else:
                        if k=="Any":
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1] in self.LSs]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1]==l]) == 0
                        else:
                            if l=="Any":
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1] in self.LSs and arc[2]==k]) == 0
                            else:
                                prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.proccap.keys() if arc[0].startswith(c) and arc[1]==l and arc[2]==k]) == 0
                    self.n_constr += 1
                except:
                    print "<<ERROR>> Could not exclude Procurement Decision:"
                    print item
                    print "    for t= " + month
                    self.errors += 1

        # Input Routing Constraints
        for route in self.user_add_route.items(): # route = (loc1, loc2, com) (mt,t)
            for month in route[1][1]:
                try:
                    t = self.horizon.index(month)
                    if route[0][1]=="Any":
                        if route[0][2]=="Any":
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==route[0][0]]) >= float(route[1][0])
                        else:
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==route[0][0] and arc[2]==route[0][2]]) >= float(route[1][0])
                    else:
                        if route[0][2]=="Any":
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==route[0][0] and arc[1]==route[0][1]]) >= float(route[1][0])
                        else:
                            prob += self.F[route[0][0],route[0][1],route[0][2],t] >= float(route[1][0])
                    self.n_constr += 1
                except :
                    print "<<ERROR>> Could not include Routing Decision:"
                    print route
                    print "    for t= " + month
                    self.errors += 1

        # Exclude Route
        for route in self.user_ex_route.items(): # route = (loc1, loc2, com) (t)
            for month in route[1]:
                try:
                    t = self.horizon.index(month)
                    if route[0][1]=="Any":
                        if route[0][2]=="Any":
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==route[0][0]]) == 0
                        else:
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==route[0][0] and arc[2]==route[0][2]]) == 0
                    else:
                        if route[0][2]=="Any":
                            prob += lpSum([self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==route[0][0] and arc[1]==route[0][1]]) == 0
                        else:
                            prob += self.F[route[0][0],route[0][1],route[0][2],t] == 0
                    self.n_constr += 1
                except :
                    print "<<ERROR>> Could not exclude Routing Decision:"
                    print route
                    print "    for t= " + month
                    self.errors += 1

        # Capacity Utilisation
        for i in self.user_cap_util.items(): # i = (loc,t) (min,max)
            if i[0][1] not in self.horizon:
                continue
            try:
                l,t,mn,mx = i[0][0],self.horizon.index(i[0][1]),float(i[1][0]),float(i[1][1])
                prob += self.LOAD[l,t] <= mx/100.0 * self.nodecap[l,i[0][1]]
                prob += self.LOAD[l,t] >= mn/100.0 * self.nodecap[l,i[0][1]]
                self.n_constr += 2
            except:
                print "<<ERROR>> Could not set Capacity Utilisation Decision:"
                print i
                self.errors += 1

        # Capacity Allocation
        for i in self.user_cap_aloc.items(): # i = (loc,t) (min,max)
            if i[0][1] not in self.horizon:
                continue
            l,mn,mx = i[0][0],float(i[1][0]),float(i[1][1])
            if l in self.DPs:
                s = list(self.DPs)
            else:
                s = list(self.EDPs)
            try:
                t = self.horizon.index(i[0][1])
                prob += lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==l) <= mx/100.0 * lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0] in s)
                prob += lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0]==l) >= mn/100.0 * lpSum(self.F[arc[0],arc[1],arc[2],t] for arc in self.cost.keys() if arc[0] in s)
                self.n_constr += 2
            except:
                print "<<ERROR>> Could not set Capacity Allocation Decision:"
                print i
                self.errors += 1

        # Include Food Group
        for item in self.user_add_fg.items(): # item = (fg) (min,max)
            fg,mn,mx=item[0],item[1][0],item[1][1]
            for t in self.hor:
                if t in self.empty:
                    continue
                try:
                    prob += self.G[fg][t] == 1
                    self.n_constr += 1
                    if mn != "N/A" and mn != "":
                        prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == fg) >= float(mn)
                        self.n_constr += 1
                    if mx != "N/A" and mx != "":
                        prob += lpSum(self.R[k][t] for k in self.commodities if self.group[k] == fg) <= float(mx)
                        self.n_constr += 1
                except:
                    print "<<ERROR>> Could not include Food Group Decision:"
                    print item
                    print "    for t= " + t
                    self.errors += 1

        # Exclude Food Group
        for g in self.user_ex_fg:
            for t in self.hor:
                prob += self.G[g][t] == 0
                self.n_constr += 1

        # Amount of unique commodities
        for t in self.hor:
            if t in self.empty:
                continue
            if self.user_add_mincom.get() != "N/A":
                prob += self.COMS[t] >= float(self.user_add_mincom.get())
                self.n_constr += 1
            if self.user_add_maxcom.get() != "N/A":
                prob += self.COMS[t] <= float(self.user_add_maxcom.get())
                self.n_constr += 1

        # Energy from fat/protein
        for t in self.hor:
            if t in self.empty:
                continue
            prob += self.PROT[t] >= float(self.user_nut_minprot.get())/100 * self.KCAL[t]
            prob += self.PROT[t] <= float(self.user_nut_maxprot.get())/100 * self.KCAL[t]
            prob += self.FAT[t] >= float(self.user_nut_minfat.get())/100 * self.KCAL[t]
            prob += self.FAT[t] <= float(self.user_nut_maxfat.get())/100 * self.KCAL[t]
            self.n_constr += 4

        # Allow shortfalls?
        if (self.allowshortfalls.get() == 0):
            for l in self.nutrients:
                for t in self.hor:
                    if t in self.empty:
                        continue
                    prob += self.S[l][t] == 0
                    self.n_constr += 1
        else:
            for item in self.user_add_nut.items(): # item = (nut) (max %)
                if item[1] != "N/A":
                    for t in self.hor:
                        if t in self.empty:
                            continue
                        if item[0] in self.nutrients:
                            prob += self.S[item[0]][t] <= float(item[1])/100
                            self.n_constr += 1
                        else: # item[0]="All"
                            for l in self.nutrients:
                                prob += self.S[l][t] <= float(item[1])/100
                                self.n_constr += 1

        # GMO
        if (self.gmo.get() == 0): # GMO not allowed
            for key in self.proccap.keys(): # key = (src, ndp, com)
                if self.isGMO[key] == 1:
                    for t in self.hor:
                        prob += self.F[key[0],key[1],key[2],t] == 0
                        self.n_constr += 1

        # Solution constraints
        for key in self.mingoal.keys(): # key = (statistic, period)
            if key[1] in self.horizon:
                t = self.horizon.index(key[1])
            elif key[1] == "Average" or key[1] == "Total":
                t = key[1]
            else:
                print "<<<ERROR>>> Could not add Statistic Constraint for period: ", key[1]
                continue
            if key[0] != "Lead Time":
                if self.mingoal[key]!="N/A" and self.mingoal[key]!="":
                    prob += self.stats[key[0]][t] >= float(self.mingoal[key])
                    self.n_constr += 1
                if self.maxgoal[key]!="N/A" and self.maxgoal[key]!="":
                    prob += self.stats[key[0]][t] <= float(self.maxgoal[key])
                    self.n_constr += 1
            else: # LT is a bit tricky, so we need some minor adjustments
                if t in self.hor:
                    if self.mingoal[key]!="N/A" and self.mingoal[key]!="":
                        prob += lpSum([self.P[item[0]][t] for item in self.quick.items() if item[1]>=float(self.mingoal[key])]) >= 1
                        prob += self.stats[key[0]][t] >= float(self.mingoal[key])
                        self.n_constr += 2
                    if self.maxgoal[key]!="N/A" and self.maxgoal[key]!="":
                        for item in self.quick.items(): # item = (src,ndp)(dur)
                            if item[1] > float(self.maxgoal[key]):
                                prob += self.P[item[0]][t] == 0
                                self.n_constr += 1
                        prob += self.stats[key[0]][t] <= float(self.maxgoal[key])
                        self.n_constr += 1
                else: # t = average or total
                    if t=="Average":
                        if self.mingoal[key]!="N/A" and self.mingoal[key]!="":
                            prob += self.LTsum["Total"] >= float(self.mingoal[key]) * self.MT["Total"]
                            self.n_constr += 1
                        if self.maxgoal[key]!="N/A" and self.maxgoal[key]!="":
                            prob += self.LTsum["Total"] <= float(self.maxgoal[key]) * self.MT["Total"]
                            self.n_constr += 1
                        # NB:   ALT = LTsum/MT  =>   ALT>=min  <=> LTsum >= min * MT
                    else: # t == "Total"
                        if self.mingoal[key]!="N/A" and self.mingoal[key]!="":
                            prob += self.LTmax >= float(self.mingoal[key])
                            self.n_constr += 1
                        if self.maxgoal[key]!="N/A" and self.maxgoal[key]!="":
                            prob += self.LTmax <= float(self.maxgoal[key])
                            self.n_constr += 1
                            for item in self.quick.items(): # item = (src,ndp)(dur)
                                if item[1] > float(self.maxgoal[key]):
                                    for t in self.hor:
                                        prob += self.P[item[0]][t] == 0
                                        self.n_constr += 1
            # NB: The model is bad at handling lead times, so we add some redundant constraints to make it easier to find solutions that adhere to the LTmax constraints



        print "Finished defining the optimisation model:"
        print str(self.n_vars) + " Variables & " + str(self.n_constr) + " Constraints"
        print " "
        print "Solving..."
        prob.writeLP("AIDM.lp") # The problem data is written to an .lp file
        prob.solve() # The problem is solved using PuLP's choice of Solver
        self.calcdur = time.time()-self.calcstart
        print "Solved!"
        self.status = LpStatus[prob.status]
        print "Solver status: ", self.status
        print "Scenario time: " , self.fmt_wcommas(self.calcdur)[1:] + " seconds"
        print " "

        if self.status == "Optimal":
            self.display_outputs(NAME) # Show KPIs for the solution
        self.frame_right.update_idletasks()

        mypath = os.path.dirname(os.path.abspath(__file__))
        (_, _, filenames) = os.walk(mypath).next()
        for f in filenames:
            if f.endswith(".mps"):
                try:
                    os.remove(f)
                except:
                    None

    def objset(self):
        '''
        Solve the current scenario.
        '''

        if self.prepped != [self.tstart.get(),self.tend.get(),self.useforecasts.get()]: # the general constraints have not been set yet for the current time horizon
            self.prep()
        # set up analysis
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'output')
        self.calculate(self.scenname.get())
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(dest_dir,self.scenname.get()),self.scenname.get())
        self.countscen+=1
        self.scenname.set("Scenario_"+str(self.countscen).zfill(3))

    def autoanalysis(self):
        '''
        Provides the user with some automated analyses to gain quick insights.
        '''

        # set up
        if self.checkreset.get() == 1:
            self.reset()
        print "Starting automated analysis..."
        print " "
        tick = time.time()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'output')
        self.quick_save()
        self.infeas = []
        self.analysis_count = 0

        # run the checked Design Analyses
        for s in self.nvs_scens:
            if self.checkbox[s].get() == 1:
                self.auto_des_nvs(dest_dir)
                break
        for s in self.cv_scens:
            if self.checkbox[s].get() == 1:
                self.auto_des_cv(dest_dir)
                break
        for s in self.lt_scens:
            if self.checkbox[s].get() == 1:
                self.auto_des_lt(dest_dir)
                break

        # run the checked Trade-Off Analyses
        if self.obj2.get() != "None":
            self.auto_to_obj(dest_dir)
        if self.breakdown.get() != "None":
            self.auto_to_src(dest_dir)

        # run the checked Adjustment Analyses
        if self.checkbox["Remove 1 Commodity"].get()==1:
            self.auto_adj_rem(dest_dir)
        if self.checkbox["Replace 1 Commodity"].get()==1:
            self.auto_adj_swap(dest_dir)
        if self.checkbox["Optimise Ration Sizes"].get()==1:
            self.auto_adj_rat(dest_dir)
        if self.checkbox["Adjust Transfer Modality"].get()==1:
            self.auto_adj_cv(dest_dir)
        if self.checkbox["Increase Prices"].get()==1:
            self.auto_adj_proc(dest_dir)
        if self.checkbox["Scale Up Operation"].get()==1:
            self.auto_adj_scaleup(dest_dir)
        if self.checkbox["Sourcing Breakdown"].get()==1:
            self.auto_adj_src(dest_dir)
        if self.checkbox["Adjust US-IK Funding"].get()==1:
            self.auto_adj_ik(dest_dir)
        if self.checkbox["Allocate Resources"].get()==1:
            print "The Allocate Resources analysis is currently on hold and will be available again in a future release."
            # self.auto_adj_shapley(dest_dir)

        # wrap up
        if len(self.infeas)>0:
            print "<<WARNING>> Some scenarios did not solve correctly:"
            for item in self.infeas:
                print "%-35s %s" % (item[0], item[1])
            print " "
        self.display_benchmarks()
        self.close(self.autowin)
        tack = time.time()
        print "Finished automated analysis"
        print "Analysed " + str(self.analysis_count) + " scenarios in " + self.fmt_wcommas(tack-tick)[1:] + " seconds"
        self.csv_benchmarks(dest_dir,"Automated Analysis")
        print " "
        self.quick_load()

    def auto_des_nvs(self, dest_dir):
        '''
        Automated analysis (design): Nutritional Value Score
        '''

        # set up
        if self.prepped != [self.tstart.get(),self.tend.get(),self.useforecasts.get()]: # the general constraints have not been set yet for the current time horizon
            self.prep()
        sub_dir = os.path.join(dest_dir, 'NVS Scenarios')
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.allowshortfalls.set(1)
        self.empty = [] # this set is usually initialised in self.calculate, but we need it beforehand here
        for t in self.hor:
            if sum(self.dem[self.ben.get(),i,self.horizon[t]] for i in self.FDPs) == 0 : # 'empty' month
                self.empty.append(t)

        # evaluate the selected scenarios
        for s in self.nvs_scens:
            if self.checkbox[s].get()==1:
                perc = float(s[7:-5])/100
                nvsmin = perc * 11
                for t in self.hor:
                    if t not in self.empty:
                        self.mingoal["Nutritional Value Score",self.horizon[t]] = nvsmin
                        self.maxgoal["Nutritional Value Score",self.horizon[t]] = 11
                self.calculate(s)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,s),s)
                else:
                    self.infeas.append([s,self.status])
                self.checkbox[s].set(0)

        # wrap up
        self.csv_benchmarks(sub_dir, "NVS Scenarios")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.quick_load()


    def auto_des_cv(self, dest_dir):
        '''
        Automated analysis (design): C&V
        '''

        # set up
        if self.prepped != [self.tstart.get(),self.tend.get(),self.useforecasts.get()]: # the general constraints have not been set yet for the current time horizon
            self.prep()
        sub_dir = os.path.join(dest_dir, 'C&V Scenarios')
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.empty = []

        # evaluate the selected scenarios
        for s in self.cv_scens:
            if self.checkbox[s].get()==1:
                perc = s[9:-5]
                for t in self.hor:
                    self.user_cv_min.set(perc)
                    self.user_cv_max.set(perc)
                self.calculate(s)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,s),s)
                else:
                    self.infeas.append([s,self.status])
                self.checkbox[s].set(0)

        # wrap up
        self.csv_benchmarks(sub_dir, "C&V Scenarios")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.quick_load()

    def auto_des_lt(self, dest_dir):
        '''
        Automated analysis (design): Lead Times
        '''

        # set up
        if self.prepped != [self.tstart.get(),self.tend.get(),self.useforecasts.get()]: # the general constraints have not been set yet for the current time horizon
            self.prep()
        sub_dir = os.path.join(dest_dir, 'LT Scenarios')
        check, val = 0, 0
        bmcopy = self.solutions.copy()
        self.solutions = {}

        # evaluate the selected LT scenarios
        for s in self.lt_scens:
            if self.checkbox[s].get()==1:
                days = float(s[0:4])
                for t in self.hor:
                    self.mingoal["Lead Time","Total"] = 0
                    self.maxgoal["Lead Time","Total"] = days
                self.calculate(s)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,s),s)
                    check, val = 1, days
                else:
                    self.infeas.append([s,self.status])
                    if check == 1:
                        print "Minimum Lead Time determined:", val
                        break # Lead Times are in descending order, so if you can't solve a scenario you also can't solve the subsequent scenarios
                self.checkbox[s].set(0)

        # wrap up
        self.csv_benchmarks(sub_dir, "LT Scenarios")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.quick_load()

    def auto_to_obj(self, dest_dir):
        '''
        Automated analysis (trade-off): Secondary objective
        '''

        # set up
        if self.prepped != [self.tstart.get(),self.tend.get(),self.useforecasts.get()]: # the general constraints have not been set yet for the current time horizon
            self.prep()
        s = self.obj2.get()
        sub_dir = os.path.join(dest_dir, "(TO) " + s)
        name = "(TO) " + s + " - "
        val = float(self.objmin.get())
        maxval = float(self.objmax.get())
        incr = float(self.increment.get())
        itnum = int(round((maxval-val)/incr) + 1)
        i = 0
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.empty = []
        for t in self.hor:
            if sum(self.dem[self.ben.get(),i,self.horizon[t]] for i in self.FDPs) == 0 : # 'empty' month
                self.empty.append(t)

        # run a trade-off analysis depending on the chosen secondary objective
        if s == "NVS (Min)" :
            check = 0
            self.allowshortfalls.set(1)
            for i in range(itnum):
                if val < 10:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                for t in self.hor:
                    if t not in self.empty:
                        self.mingoal["Nutritional Value Score",self.horizon[t]] = val
                        self.maxgoal["Nutritional Value Score",self.horizon[t]] = maxval
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                    check = 1
                else:
                    self.infeas.append([name2,self.status])
                    if check == 1:
                        print "Maximum NVS determined: ", val-incr
                        break
                val += incr
        elif s == "NVS (Avg)":
            check = 0
            self.allowshortfalls.set(1)
            for i in range(itnum):
                if val < 10:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                self.mingoal["Nutritional Value Score","Average"] = val
                self.maxgoal["Nutritional Value Score","Average"] = val
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                    check = 1
                else:
                    self.infeas.append([name2,self.status])
                    if check == 1:
                        print "Maximum NVS determined: ", val-incr
                        break
                val += incr
        elif s == "NVS (% Supplied)":
            check = 0
            self.allowshortfalls.set(1)
            for i in range(itnum):
                if val < 10:
                    name2 = name + "    " + str(val)
                elif val < 100:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                self.mingoal["Nutritional Value Score","Average"] = val/100.0*11
                self.maxgoal["Nutritional Value Score","Average"] = val/100.0*11
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                    check = 1
                else:
                    self.infeas.append([name2,self.status])
                    if check == 1:
                        print "Maximum NVS determined: ", val-incr
                        break
                val += incr
        elif s == "Kcal (Avg)":
            check = 0
            self.allowshortfalls.set(1)
            for i in range(itnum):
                if val < 10:
                    name2 = name + "      " + str(val)
                elif val < 100:
                    name2 = name + "    " + str(val)
                elif val < 1000:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                self.mingoal["Energy Supplied [Total]","Average"] = val
                self.maxgoal["Energy Supplied [Total]","Average"] = val
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                    check = 1
                else:
                    self.infeas.append([name2,self.status])
                    if check == 1:
                        print "Maximum kcal determined: ", val-incr
                        break
                val += incr
        elif s == "C&V (%)":
            check = 0
            for i in range(itnum):
                if val < 10:
                    name2 = name + "    " + str(val)
                elif val < 100:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                self.user_cv_min.set(str(val))
                self.user_cv_max.set(str(val))
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                    check = 1
                else:
                    self.infeas.append([name2,self.status])
                    if check == 1:
                        print "Maximum % C&V determined:", val-incr
                        break # If x% is feasible but x+incr% is not, y>x+incr will also be infeasible
                val += incr
        elif s == "Loc (%)":
            check = 0
            for i in range(itnum):
                if val < 10:
                    name2 = name + "    " + str(val)
                elif val < 100:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                self.user_loc_min.set(str(val))
                self.user_loc_max.set(str(val))
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                    check = 1
                else:
                    self.infeas.append([name2,self.status])
                    if check == 1:
                        print "Maximum % Loc determined:", val-incr
                        break # If x% is feasible but x+incr% is not, y>x+incr will also be infeasible
                val += incr
        elif s == "Lead Time (Avg)":
            check = 0
            for i in range(itnum):
                if val < 10:
                    name2 = name + "    " + str(val)
                elif val < 100:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                self.mingoal["Lead Time","Average"] = val
                self.maxgoal["Lead Time","Average"] = val
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                    check = 1
                else:
                    self.infeas.append([name2,self.status])
                    if check == 1:
                        print "Maximum Lead Time determined: ", val-incr
                        break
                val += incr
        elif s == "Lead Time (Max)":
            for i in range(itnum):
                if val < 10:
                    name2 = name + "    " + str(val)
                elif val < 100:
                    name2 = name + "  " + str(val)
                else:
                    name2 = name + str(val)
                self.mingoal["Lead Time","Total"] = val
                self.maxgoal["Lead Time","Total"] = val
                self.calculate(name2)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name2),name2)
                else:
                    self.infeas.append([name2,self.status])
                val += incr

        # wrap up
        self.quick_load()
        self.csv_benchmarks(sub_dir,name[:-3])
        self.obj2.set("None")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()

    def auto_to_src(self, dest_dir):
        '''
        Automated analysis (trade-off): Sourcing breakdown
        '''

        # set up
        g = self.breakdown.get()
        sub_dir = os.path.join(dest_dir, "(TO) Sourcing (" + g + ")")
        bmcopy = self.solutions.copy()
        bmcopy2 = {}

        if g != "All":
            # find all relevant commodities and sources
            self.solutions = {}
            coms = []
            for k in self.commodities:
                if self.group[k] == g:
                    coms.append(k)
            self.reset()
            self.allowshortfalls.set(1)
            self.sensible.set(0)
            self.activities = []
            opt = {}
            for k in coms:
                opt[k] = []
            for key in self.proccap.keys():
                if key[2] in coms:
                    opt[key[2]].append(key)
            # evaluate all commodities
            for k in opt.keys():
                # find the optimal sourcing plan (may utilise multiple sources for this commodity)
                self.reset_fix()
                self.food2fix.append((k,"100"))
                self.calculate("Optimal sourcing for " + k)
                self.analysis_count += 1
                best = value(self.TC["Total"])
                self.csv_outputs("\\\\?\\" + os.path.join(sub_dir,"Optimal sourcing for " + k),"Optimal sourcing for " + k)
                check = 0
                # evaluate all single-source options for this commodity
                for proc in opt[k]: # proc = (src, ndp, com)
                    if proc[1] not in self.LMs:
                        i = proc[1].index("(")-1
                        name = proc[2] + " from " + proc[0][:-6] + " " + proc[0][-3:] + " " + proc[1][:i]
                    else:
                        if check == 0: # check whether C&V has been evaluated yet
                            name = proc[2] +" from " + proc[0][:-6] + " " + proc[0][-3:]
                            check = 1
                        else:
                            continue
                    self.reset_proc()
                    self.reset_cv()
                    # exclude all other sourcing options
                    for arc in opt[k]: # arc = (src, ndp, com)
                        if arc != proc:
                            if arc[1] in (self.ISs + self.RSs):
                                self.user_ex_proc_int[arc] = self.horizon
                            elif arc[1] in self.LSs:
                                self.user_ex_proc_loc[arc] = self.horizon
                            else: # in LMs
                                if proc[1] not in self.LMs:
                                    self.user_ex_cv[arc[1],arc[2]] = self.horizon
                    self.calculate(name)
                    self.analysis_count += 1
                    if self.status == "Optimal":
                        self.csv_outputs("\\\\?\\" + os.path.join(sub_dir,name),name)
                        if abs(best - value(self.TC["Total"])) < 10:
                            self.solutions.pop("Optimal sourcing for " + k,None) # Only show optimal solution in the output when it does not overlap with a single-source solution
                    else:
                        self.infeas.append([name,self.status])
                        print " "
            # wrap up
            self.csv_benchmarks(sub_dir, "(TO) Sourcing (" + g + ")")
            bmcopy.update(self.solutions)
            self.solutions = bmcopy.copy()
        else: # Run analysis for all food groups
            for fg in self.foodgroups:
                self.solutions = {}
                subsub_dir = os.path.join(sub_dir,fg)
                coms = []
                for k in self.commodities:
                    if self.group[k] == fg:
                        coms.append(k)
                self.reset()
                self.allowshortfalls.set(1)
                self.sensible.set(0)
                self.activities = []
                opt = {}
                for k in coms:
                    opt[k] = []
                for key in self.proccap.keys():
                    if key[2] in coms:
                        opt[key[2]].append(key)
                for k in opt.keys():
                    self.reset_fix()
                    self.food2fix.append((k,"100"))
                    self.calculate("Optimal sourcing for " + k)
                    self.analysis_count += 1
                    best = value(self.TC["Total"])
                    self.csv_outputs("\\\\?\\" + os.path.join(subsub_dir,"Optimal sourcing for " + k),"Optimal sourcing for " + k)
                    check = 0
                    for proc in opt[k]:
                        if proc[1] not in self.LMs:
                            i = proc[1].index("(")
                            name = proc[2] + " from " + proc[0][:-6] + " " + proc[0][-3:] + " " + proc[1][:i]
                        else:
                            if check == 0:
                                name = proc[2] +" from " + proc[0][:-6] + " " + proc[0][-3:]
                                check = 1
                            else:
                                continue
                        self.reset_proc()
                        self.reset_cv()
                        for arc in opt[k]:
                            if arc != proc:
                                if arc[1] in (self.ISs + self.RSs):
                                    self.user_ex_proc_int[arc] = self.horizon
                                elif arc[1] in self.LSs:
                                    self.user_ex_proc_loc[arc] = self.horizon
                                else: # in LMs
                                    if proc[1] not in self.LMs:
                                        self.user_ex_cv[arc[1],arc[2]] = self.horizon
                        self.calculate(name)
                        self.analysis_count += 1
                        if self.status == "Optimal":
                            self.csv_outputs("\\\\?\\" + os.path.join(subsub_dir,name),name)
                            if abs(value(self.TC["Total"]) - best) < 10:
                                self.solutions.pop("Optimal sourcing for " + k,None) # Only show optimal solution in the output when it does not overlap with a single-source solution
                        else:
                            self.infeas.append([name,self.status])
                            print " "
                # wrap up the food group
                self.csv_benchmarks(subsub_dir, "(TO) Sourcing (" + fg + ")")
                bmcopy2.update(self.solutions)
                self.solutions = bmcopy2.copy()
            # wrap up the 'All' analysis
            self.csv_benchmarks(sub_dir, "(TO) Sourcing (All)")
            bmcopy.update(self.solutions)
            self.solutions = bmcopy.copy()

        # wrap up
        self.breakdown.set("None")
        self.quick_load()

    def auto_adj_rem(self, dest_dir):
        '''
        Automated analysis (adjustment): Remove 1
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Remove 1 Commodity (" + self.baseline.get() +")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # evaluate the baseline scenario and set up analysis
        name = self.baseline.get()
        self.calculate(name)
        self.analysis_count += 1
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)
        self.allowshortfalls.set(1)
        self.sensible.set(0)
        comtoremove = []
        for item in self.food2fix: # (com, quantity)
            if item[1] != "0":
                comtoremove.append(item[0])

        # remove each commodity in the food basket
        for k in comtoremove:
            self.remove = k
            name = "Remove " + k
            self.calculate(name)
            self.analysis_count += 1
            if self.status == "Optimal":
                self.csv_outputs(os.path.join(sub_dir,name),name)
            else:
                self.infeas.append([name,self.status])

        # wrap up
        self.csv_benchmarks(sub_dir, "Remove 1 Commodity")
        self.remove = ""
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Remove 1 Commodity"].set(0)

    def auto_adj_swap(self, dest_dir):
        '''
        Automated analysis (adjustment): Swap 1
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Replace 1 Commodity (" + self.baseline.get() +")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # evaluate the baseline scenario and set up analysis
        name = self.baseline.get()
        self.calculate(name)
        self.analysis_count += 1
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)
        self.allowshortfalls.set(1)
        self.sensible.set(0)
        inbasket = {}
        outbasket = {}
        for g in self.foodgroups:
            inbasket[g]=[]
            outbasket[g]=[]
        for item in self.food2fix: # (com, quantity)
            if item[1] != "0":
                inbasket[self.group[item[0]]].append(item[0])
        for k in self.commodities:
            gr = self.group[k]
            if k not in inbasket[gr] and k != "CASH":
                outbasket[gr].append(k)

        # evaluate all possible swaps
        for g in self.foodgroups:
            for k1 in inbasket[g]:
                for k2 in outbasket[g]:
                    self.replace = [k1,k2]
                    name = "Swap " + k1 + " with " + k2
                    self.calculate(name)
                    self.analysis_count += 1
                    if self.status == "Optimal":
                        self.csv_outputs(os.path.join(sub_dir,name),name)
                    else:
                        self.infeas.append([name,self.status])

        # wrap up
        self.csv_benchmarks(sub_dir, "Replace 1 Commodity")
        self.replace = []
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Replace 1 Commodity"].set(0)

    def auto_adj_rat(self, dest_dir):
        '''
        Automated analysis (adjustment): Optimise ration sizes
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Optimise Ration Sizes (" + self.baseline.get() +", " + self.deviation.get() + ")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # evaluate the baseline scenario and set up analysis
        name = self.baseline.get()
        self.calculate(name)
        self.analysis_count += 1
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)
        d = self.deviation.get()
        if d.endswith("%"):
            self.ration = float(d[:-1])/100
        else:
            self.ration = float(d)/100
        for t in self.hor:
            self.mingoal["Commodities (#)",self.horizon[t]]=value(self.COMS[t])
            self.maxgoal["Commodities (#)",self.horizon[t]]=value(self.COMS[t])
        self.allowshortfalls.set(1)
        self.sensible.set(0)
        check = 0

        # evaluate NVS levels of 5.5-11, with increments of .5
        nvs = 5.5
        for i in range(12):
            for t in self.hor:
                if t not in self.empty:
                    self.mingoal["Nutritional Value Score",self.horizon[t]] = nvs + .5*i
                    self.maxgoal["Nutritional Value Score",self.horizon[t]] = nvs + .5*i
            if nvs + .5*i < 10:
                name = "Optimised ration for   " + str(nvs + .5*i) + " NVS"
            else:
                name = "Optimised ration for " + str(nvs + .5*i) + " NVS"
            self.calculate(name)
            self.analysis_count += 1
            if self.status == "Optimal":
                self.csv_outputs(os.path.join(sub_dir,name),name)
                check = 1
            else:
                self.infeas.append([name,self.status])
                if check == 1:
                    print "Maximum NVS determined:", nvs + (i-1)*.5
                    break # If X NVS is feasible but X+0.5 is not, Y>X+0.5 will also be infeasible

        # wrap up
        self.csv_benchmarks(sub_dir,"Optimise Ration Sizes")
        self.ration = ""
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Optimise Ration Sizes"].set(0)

    def auto_adj_ik(self, dest_dir):
        '''
        Automated analysis (adjustment): Analyse In-Kind funding
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Adjust US-IK Funding (" + self.baseline.get() + ")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # evaluate the baseline scenario and set up analysis
        name = self.baseline.get()
        self.calculate(name)
        self.analysis_count += 1
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)
        self.reset_ik()
        check = 0

        # evaluate 0-100% C&V ratio, with increments of 5%
        for i in range(21):
            if i < 2:
                l = "    "
            elif i < 20:
                l = "  "
            else:
                l = ""
            name = "US-IK Funding " + l + str(i*5) + "%"
            self.user_add_ik["USD","Percentage",self.horizon[0]] = i*5
            self.calculate(name)
            self.analysis_count += 1
            if self.status == "Optimal":
                self.csv_outputs(os.path.join(sub_dir,name),name)
                check = 1
            else:
                self.infeas.append([name,self.status])
                if check == 1:
                    print "Maximum % US-IK determined:",(i-1)*5
                    break # If x% is feasible but x+5% is not, y>x+5 will also be infeasible

        # wrap up
        self.csv_benchmarks(sub_dir,"Adjust US-IK Funding")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Adjust US-IK Funding"].set(0)

    def auto_adj_proc(self, dest_dir):
        '''
        Automated analysis (adjustment): Analyse increases in local/regional procurement prices
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Increase Prices (" + self.baseline.get() + ")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # evaluate the baseline scenario and set up analysis
        name = self.baseline.get()
        self.calculate(name)
        self.analysis_count += 1
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)
        self.reset_proc()

        # evaluate -50 to +100% loc/reg price increases, with increments of 10%
        for i in range(16):
            self.prepped = ["","",""]
            l = ""
            if self.incr_loc.get()==1:
                l += "loc "
                self.mod_loc = .5+i*.1
            if self.incr_reg.get()==1:
                l += "reg "
                self.mod_reg = .5+i*.1
            if self.incr_cbt.get()==1:
                l += "cbt "
                self.mod_cbt = .5+i*.1
            if i == 5:
                l += "      "
            elif i == 15:
                l += ""
            else:
                l += "  "
            if i > 5:
                l += "+"
            else:
                l += " "
            name = "Price " + l + str(-50+i*10) + "%"
            self.calculate(name)
            self.analysis_count += 1
            if self.status == "Optimal":
                self.csv_outputs(os.path.join(sub_dir,name),name)
                check = 1
            else:
                self.infeas.append([name,self.status])

        # wrap up
        self.csv_benchmarks(sub_dir,"Increase Prices")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Increase Prices"].set(0)
        self.incr_loc.set(0)
        self.incr_reg.set(0)
        self.mod_loc, self.mod_reg, self.mod_cbt = 1, 1, 1

    def auto_adj_cv(self, dest_dir):
        '''
        Automated analysis (adjustment): Analyse C&V ratio
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Adjust Transfer Modality (" + self.baseline.get() + ")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # evaluate the baseline scenario and set up analysis
        name = self.baseline.get()
        self.calculate(name)
        self.analysis_count += 1
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)
        self.totalmt = value(self.MT["Total"]) # if we don't specify the mt the tool will start buying more than we need just to get a high %-rate
        self.reset_proc()
        self.reset_cv()
        check = 0

        # evaluate 0-100% C&V ratio, with increments of 5%
        for i in range(21):
            if i < 2:
                l = "    "
            elif i < 20:
                l = "  "
            else:
                l = ""
            name = "Adjusted Transfer Modality " + l + str(i*5) + "%"
            self.user_cv_min.set(str(i*5))
            self.user_cv_max.set(str(i*5))
            self.calculate(name)
            self.analysis_count += 1
            if self.status == "Optimal":
                self.csv_outputs(os.path.join(sub_dir,name),name)
                check = 1
            else:
                self.infeas.append([name,self.status])
                if check == 1:
                    print "Maximum % C&V determined:",(i-1)*5
                    break # If x% is feasible but x+5% is not, y>x+5 will also be infeasible

        # wrap up
        self.csv_benchmarks(sub_dir,"Adjust Transfer Modality")
        self.totalmt = ""
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Adjust Transfer Modality"].set(0)

    def auto_adj_scaleup(self, dest_dir):
        '''
        Automated analysis (adjustment): Analyse demand increases
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Scale Up Operation (" + self.baseline.get() + ")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # evaluate the baseline scenario and set up analysis
        name = self.baseline.get()
        self.calculate(name)
        self.analysis_count += 1
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)
        check = 0

        # investigate demand increases of 5-100%
        for i in range(1,21):
            if i == 1:
                l = "    "
            elif i < 20:
                l = "  "
            else:
                l = ""
            name = "Increase Demand By " + l + str(i*5) + "%"
            self.scaleup = 1 + i*5/100.0
            self.calculate(name)
            self.analysis_count += 1
            if self.status == "Optimal":
                self.csv_outputs(os.path.join(sub_dir,name),name)
                check = 1
            else:
                self.infeas.append([name,self.status])
                if check == 1:
                    print "Maximum % increase in demand determined:",(i-1)*5
                    break # If x% is feasible but x+5% is not, y>x+5 will also be infeasible

        # wrap up
        self.csv_benchmarks(sub_dir, "Scale Up Operation")
        self.scaleup = 1
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Scale Up Operation"].set(0)

    def auto_adj_src(self, dest_dir):
        '''
        Automated analysis (adjustment): Sourcing breakdown for current food basket
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Sourcing Breakdown (" + self.baseline.get() + ")")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # find all sources for all commodities in the current basket
        curbasket = []
        for item in self.food2fix: # (com, quantity)
            curbasket.append(item[0])
        self.reset()
        self.allowshortfalls.set(1)
        self.sensible.set(0)
        self.activities = []
        opt = {}
        for k in curbasket:
            opt[k] = []
        for key in self.proccap.keys():
            if key[2] in curbasket:
                opt[key[2]].append(key)

        # calculate the cost of all sources
        for k in opt.keys():
            self.reset_fix()
            self.food2fix.append((k,"100"))
            # first we find the optimal sourcing strategy (may utilise multiple sources for one commodity)
            self.calculate("Optimal sourcing for " + k)
            self.analysis_count += 1
            best = value(self.TC["Total"])
            self.csv_outputs("\\\\?\\" + os.path.join(sub_dir,"Optimal sourcing for " + k),"Optimal sourcing for " + k) # the file names get ridiculously long, by adding "\\\\?\\" we can circumvent namelength errors
            check = 0
            # then we investigate each single-source strategy by excluding all other sources
            for proc in opt[k]:
                if proc[1] not in self.LMs:
                    i = proc[1].index("(")
                    name = proc[2] + " from " + proc[0][:-6] + " " + proc[0][-3:] + " " + proc[1][:i]
                else:
                    if check == 0: # check whether we've already investigated C&V as a source
                        name = proc[2] +" from " + proc[0][:-6] + " " + proc[0][-3:]
                        check = 1
                    else:
                        continue
                self.reset_proc()
                self.reset_cv()
                for arc in opt[k]:
                    if arc != proc:
                        if arc[1] in (self.ISs + self.RSs):
                            self.user_ex_proc_int[arc] = self.horizon
                        elif arc[1] in self.LSs:
                            self.user_ex_proc_loc[arc] = self.horizon
                        else: # in LMs
                            if proc[1] not in self.LMs:
                                self.user_ex_cv[arc[1],arc[2]] = self.horizon
                self.calculate(name)
                self.analysis_count += 1
                if self.status == "Optimal":
                    self.csv_outputs("\\\\?\\" + os.path.join(sub_dir,name),name)
                    if abs(best - value(self.TC["Total"])) < 10:
                        self.solutions.pop("Optimal sourcing for " + k,None) # Only show optimal solution in the output when it does not overlap with a single-source solution
                else:
                    self.infeas.append([name,self.status])
                    print " "

        # wrap up
        self.csv_benchmarks(sub_dir, "Sourcing Breakdown")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Sourcing Breakdown"].set(0)

    def auto_adj_shapley(self, dest_dir):
        '''
        Automated analysis (adjustment): Cost allocation using Shapley values
        '''

        # set up
        sub_dir = os.path.join(dest_dir, "Allocate Resources")
        bmcopy = self.solutions.copy()
        self.solutions = {}
        self.csvnamel.set(self.baseline.get())
        self.csv_load()

        # create all subsets
        l = self.activities
        l.append(self.ben.get())
        l = list(l)
        sets = list(chain.from_iterable( combinations(l,n) for n in range(0,len(l)+1) ))

        # create fake activity
        for t in self.periods:
            if sum(self.dem[b,t] for b in self.beneficiaries) > 0:
                for i in self.FDPs:
                    self.dem["empty",i,t] = 1
                self.dem["empty",t] = len(self.FDPs)
            else:
                for i in self.FDPs:
                    self.dem["empty",i,t] = 0
                self.dem["empty",t] = 0
        for l in self.nutrients:
            self.nutreq["empty",l] = 1
        for k in self.commodities:
            self.feedingdays["empty",k] = 1
        self.ben.set("empty")
        self.allowshortfalls.set(1)
        self.sensible.set(0)
        # NB: may need some revising, the Generate Output file doesn't like empty food baskets

        # solve for all subsets
        for s in sets:
            name = "Optimise subset " + str(s)
            self.activities = []
            for ss in s:
                self.activities.append(ss)
            self.calculate(name)
            self.analysis_count += 1
            if self.status == "Optimal":
                self.csv_outputs(os.path.join(sub_dir,name),name)
            else:
                self.infeas.append([name,self.status])

        # wrap up
        self.csv_benchmarks(sub_dir,"Allocate Resources")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.checkbox["Allocate Resources"].set(0)

    def auto_temp(self):
        '''
        Automated analysis - non-standard / temporary
        '''

        # set up
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'output')
        sub_dir = os.path.join(dest_dir, "PRRO Analysis")
        name = "April Basket"
        self.csvnamel.set(name)
        self.csv_load()
        self.calculate(name)
        if self.status == "Optimal":
            self.csv_outputs(os.path.join(sub_dir,name),name)

        # create all subsets
        l = ["WHOLE RED LENTILS","SPLIT RED LENTILS","HORSE BEANS","CHICKPEAS","WHITE BEANS"]
        abr = {}
        abr["WHOLE RED LENTILS"]="WRL"
        abr["SPLIT RED LENTILS"]="SRL"
        abr["HORSE BEANS"]="HOB"
        abr["CHICKPEAS"]="CHK"
        abr["WHITE BEANS"]="WIB"
        l = list(l)
        sets = combinations(l,3)
        for s in sets:
            for b in [0,1]:
                self.reset_fix()
                name = str(5+b) + "_ "
                self.food2fix.append(("WHEAT FLOUR",100))
                self.food2fix.append(("BULGUR WHEAT",66.67))
                self.food2fix.append(("5% BROKEN RICE",66.67))
                self.food2fix.append(("SUNFLOWER OIL",36.4))
                self.food2fix.append(("WHITE SUGAR",33.33))
                self.food2fix.append(("IODISED SALT",6.67))
                for i in s:
                    self.food2fix.append((i,33.33+b*6.67))
                    name+= " " + abr[i]
                self.calculate(name)
                if self.status == "Optimal":
                    self.csv_outputs(os.path.join(sub_dir,name),name)
                else:
                    self.infeas.append([name,self.status])

        # wrap up
        self.csv_benchmarks(sub_dir,"PRRO Analysis")

    def listanalysis(self):
        '''
        Runs a series of user-defined scenarios sequentially
        '''

        l = []
        print "Starting scenario analysis from list:"
        for i in self.listbox_scen.curselection():
            l.append(self.listbox_scen.get(i))
            print "> " + self.listbox_scen.get(i)
        if len(self.listbox_scen.curselection()) == 1:
            l.append(" ")
        tick = time.time()
        self.quick_save()
        bmcopy = self.solutions.copy()
        self.solutions = {}
        n = 0
        e = 0
        for s in l:
            if s == " ":
                continue  # NB: the for-loop goes wrong if only 1 scenario is selected, so we 'add' an empty scenario
            try:
                print "Processing scenario: " + s
                self.csvnamel.set(s)
                self.csv_load()
                self.scenname.set(s)
                self.objset()
                n += 1
            except:
                print "<<<ERROR>>> Failed to load scenario " + s
                e += 1

        self.display_benchmarks()
        tack = time.time()
        print "Analysed " + str(n) + " scenarios in " + self.fmt_wcommas(tack-tick)[1:] + " seconds"
        if e > 1:
            print "<<<WARNING>>> Encountered " + str(e) + " errors during List Analysis"
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'output')
        self.csv_benchmarks(dest_dir,"List Analysis")
        bmcopy.update(self.solutions) # Adds items from self.solutions to bmcopy. In case scenario names overlap, those from self.solutions overwrite older scenarios
        self.solutions = bmcopy.copy()
        self.quick_load()
        self.close(self.scenwin)

    def include_com(self):
        '''
        User-added constraint: Include commodity
        '''

        k = self.fb_add_speccom.get()
        k0 = self.fb_add_com.get()
        mn = self.fb_add_min.get()
        mx = self.fb_add_max.get()
        if k == "Select":
            tkMessageBox.showwarning("Error", "Must select a commodity!")
        else:
            if k == "Any":
                self.user_add_com[k0]=[mn,mx]
                print "You have added "+k0+" to the Food Basket (Min= "+mn+", Max= "+mx+")."
            else:
                self.user_add_com[k]=[mn,mx]
                print "You have added "+k+" to the Food Basket (Min= "+mn+", Max= "+mx+")."
            self.fb_add_min.set("N/A")
            self.fb_add_max.set("N/A")
            self.fb_add_com.set("Filter")
            self.fb_add_speccom.set("Select")
        print " "

    def include_proc_int(self):
        '''
        User-added constraint: Include international (or regional) purchase
        '''

        c = self.proc_add_src_int.get()
        i = self.proc_add_inco_int.get()
        k = self.proc_add_com_int.get()
        l = self.proc_add_ndp_int.get()
        q = self.proc_add_mt_int.get()
        if c == "Select" or i == "Select" or k == "Select" or l == "Select" or q == "N/A":
            tkMessageBox.showwarning("Error", "Insufficient input!")
        else:
            indices = self.listboxP.curselection()
            t = []
            for ind in indices:
                t.append(self.listboxP.get(ind))
            if t == []:
                t = list(self.horizon)
            self.user_add_proc_int[c,i,l,k]=[q,t]
            print "Fixed procurement decision (country, incoterm, ndp, commodity, mt, t): "
            print (c,i,l,k,q,t)
            self.proc_add_src_int.set("Select")
            self.proc_add_inco_int.set("Select")
            self.proc_add_ndp_int.set("Select")
            self.proc_add_com_int.set("Select")
            self.proc_add_mt_int.set("N/A")
        print " "

    def include_proc_loc(self):
        '''
        User-added constraint: Include local purchase
        '''

        c = self.proc_add_src_loc.get()
        i = self.proc_add_inco_loc.get()
        k = self.proc_add_com_loc.get()
        l = self.proc_add_ndp_loc.get()
        q = self.proc_add_mt_loc.get()
        if c == "Select" or i == "Select" or k == "Select" or l == "Select" or q == "N/A":
            tkMessageBox.showwarning("Error", "Insufficient input!")
        else:
            indices = self.listboxP.curselection()
            t = []
            for ind in indices:
                t.append(self.listboxP.get(ind))
            if t == []:
                t = list(self.horizon)
            self.user_add_proc_loc[c,i,l,k]=[q,t]
            print "Fixed procurement decision (country, incoterm, ndp, commodity, mt, t): "
            print (c,i,l,k,q,t)
            self.proc_add_src_loc.set("Select")
            self.proc_add_inco_loc.set("Select")
            self.proc_add_ndp_loc.set("Select")
            self.proc_add_com_loc.set("Select")
            self.proc_add_mt_loc.set("N/A")
        print " "

    def include_route(self):
        '''
        User-added constraint: Include supply chain leg
        '''

        o = self.route_add_loc1.get()
        d = self.route_add_loc2.get()
        k = self.route_add_com.get()
        q = self.route_add_mt.get()
        if o=="Select" or d=="Select" or k=="Select" or q=="N/A":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxR.curselection()
            t = []
            for i in indices:
                t.append(self.listboxR.get(i))
            if t == []:
                t = list(self.horizon)
            self.user_add_route[o,d,k]=[q,t]
            print "Added routing decision (Loc1, Loc2, com, mt, t):"
            print (o,d,k,q,t)
            self.route_add_loc1.set("Select")
            self.route_add_loc2.set("Select")
            self.route_add_com.set("Select")
            self.route_add_mt.set("N/A")
        print " "

    def include_fg(self):
        '''
        User-added constraint: Include food group
        '''

        fg = self.fb_add_fg.get()
        mn = self.fb_add_fg_min.get()
        mx = self.fb_add_fg_max.get()
        if fg == "Select":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            self.user_add_fg[fg]=[mn,mx]
            print "Added food group decision (fg, min, max):"
            print (fg,mn,mx)
            self.fb_add_fg.set("Select")
            self.fb_add_fg_min.set("N/A")
            self.fb_add_fg_max.set("N/A")
        print " "

    def include_cv(self):
        '''
        User-added constraint: Include C&V purchase
        '''

        m = self.cv_add_src.get()
        k = self.cv_add_com.get()
        q = self.cv_add_mt.get()
        if m=="Select" or k=="Select" or q=="N/A":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxCV.curselection()
            time = []
            for i in indices:
                time.append(self.listboxCV.get(i))
            if time == []:
                time = list(self.horizon)
            for t in time:
                self.user_add_cv[m,k,t]=q
            print "Included purchase decision (LM, com, mt, t):"
            print (m,k,q, time)
            self.cv_add_src.set("Select")
            self.cv_add_com.set("Select")
            self.cv_add_mt.set("N/A")
        print " "

    def include_ik(self):
        '''
        User-added constraint: Include In-Kind donation
        '''

        val = self.ik_donation.get()
        met = self.ik_metric.get()
        mea = self.ik_measure.get()
        if val=="N/A":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxIK.curselection()
            time = []
            for i in indices:
                time.append(self.listboxIK.get(i))
            if time == []:
                time = list(self.horizon)
            for t in time:
                self.user_add_ik[met,mea,t] = float(val)
            print "Included IK donation:"
            print (val,met,mea,t)
            self.ik_donation.set("N/A")
        print " "

    def exclude_com(self):
        '''
        User-added constraint: Exclude commodity
        '''

        k = self.fb_ex_speccom.get()
        k0 = self.fb_ex_com.get()
        if k == "Select":
            tkMessageBox.showwarning("Error", "Must select commodity!")
        else:
            if k == "Any":
                self.user_ex_com.append(k0)
                print "You have excluded "+k0+" from the Food Basket."
            else:
                self.user_ex_com.append(k)
                print "You have excluded "+k+" from the Food Basket."
            self.fb_ex_com.set("Filter")
            self.fb_ex_speccom.set("Select")
        print " "

    def exclude_proc_int(self):
        '''
        User-added constraint: Exclude international purchase
        '''

        c = self.proc_ex_src_int.get()
        l = self.proc_ex_ndp_int.get()
        k = self.proc_ex_com_int.get()
        if c == "Select" or l == "Select" or k == "Select":
            tkMessageBox.showwarning("Error", "Insufficient input!")
        else:
            indices = self.listboxP.curselection()
            t = []
            for i in indices:
                t.append(self.listboxP.get(i))
            if t == []:
                t = list(self.horizon)
            self.user_ex_proc_int[c,l,k] = t
            print "You have excluded (Source, NDP, Commodity, t):"
            print (c,l,k, t )
            self.proc_ex_src_int.set("Select")
            self.proc_ex_ndp_int.set("Select")
            self.proc_ex_com_int.set("Select")
        print " "

    def exclude_proc_loc(self):
        '''
        User-added constraint: Exclude local purchase
        '''

        c = self.proc_ex_src_loc.get()
        l = self.proc_ex_ndp_loc.get()
        k = self.proc_ex_com_loc.get()
        if c == "Select" or l == "Select" or k == "Select":
            tkMessageBox.showwarning("Error", "Insufficient input!")
        else:
            indices = self.listboxP.curselection()
            t = []
            for i in indices:
                t.append(self.listboxP.get(i))
            if t == []:
                t = list(self.horizon)
            self.user_ex_proc_loc[c,l,k] = t
            print "You have excluded (Source, NDP, Commodity, t):"
            print (c,l,k, t )
            self.proc_ex_src_loc.set("Select")
            self.proc_ex_ndp_loc.set("Select")
            self.proc_ex_com_loc.set("Select")
        print " "

    def exclude_route(self):
        '''
        User-added constraint: Exclude route
        '''

        o = self.route_ex_loc1.get()
        d = self.route_ex_loc2.get()
        k = self.route_ex_com.get()
        if o=="Select" or d=="Select" or k=="Select":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxR.curselection()
            t = []
            for i in indices:
                t.append(self.listboxR.get(i))
            if t == []:
                t = list(self.horizon)
            self.user_ex_route[o,d,k]=t
            print "Excluded routing decision (Loc1, Loc2, com, t):"
            print (o,d,k,t)
            self.route_ex_loc1.set("Select")
            self.route_ex_loc2.set("Select")
            self.route_ex_com.set("Select")
        print " "

    def exclude_fg(self):
        '''
        User-added constraint: Exclude food group
        '''

        fg = self.fb_ex_fg.get()
        if fg == "Select" :
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            self.user_ex_fg.append(fg)
            print "Excluded food group from basket:",fg
            self.fb_ex_fg.set("Select")
        print " "

    def exclude_cv(self):
        '''
        User-added constraint: Exclude C&V purchase
        '''

        m = self.cv_ex_src.get()
        k = self.cv_ex_com.get()
        if m=="Select" or k=="Select":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxCV.curselection()
            t = []
            for i in indices:
                t.append(self.listboxCV.get(i))
            if t == []:
                t = list(self.horizon)
            self.user_ex_cv[m,k] = t
            print "Excluded purchase decision (LM, com, t):"
            print (m,k,t)
            self.cv_ex_src.set("Select")
            self.cv_ex_com.set("Select")
        print " "

    def set_act(self):
        '''
        User-added constraint: Set activities to be supplied
        '''

        self.activities = []
        print "Including demand for activities:"
        for i in self.listbox_act.curselection():
            self.activities.append(self.listbox_act.get(i))
            print "> " + self.listbox_act.get(i)
        print " "
        for j in range(self.listbox_act.size()):
            if str(j) in self.listbox_act.curselection():
                self.listbox_act.itemconfig(j, background="light sky blue")
            else:
                self.listbox_act.itemconfig(j, background="white")
        self.act_button.configure(text="Select (" + str(len(self.activities)) + ")")

    def set_cap_util(self):
        '''
        User-added constraint: Set capacity utilisation
        '''

        l = self.route_util_loc.get()
        mn = self.route_util_min.get()
        mx = self.route_util_max.get()
        if l=="Select":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxR.curselection()
            time = []
            for i in indices:
                time.append(self.listboxR.get(i))
            if time == []:
                time = list(self.horizon)
            for t in time:
                self.user_cap_util[l,t]=[mn,mx]
            print "Capacity utilisation for " + l + " set to (" + mn + ", " + mx + ") for t=",time
        print " "

    def set_cap_aloc(self):
        '''
        User-added constraint: Set capacity allocation
        '''

        l = self.route_aloc_loc.get()
        mn = self.route_aloc_min.get()
        mx = self.route_aloc_max.get()
        if l=="Select":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxR.curselection()
            time = []
            for i in indices:
                time.append(self.listboxR.get(i))
            if time == []:
                time = list(self.horizon)
            for t in time:
                self.user_cap_aloc[l,t]=[mn,mx]
            print "Capacity allocation for " + l + " set to (" + mn + ", " + mx + ") for t=",time
        print " "

    def set_fb(self):
        '''
        User-added constraint: Set the food basket
        '''

        countfix=0
        self.food2fix=[]
        for i in range(0,15):
            if self.fix_com[i].get() != "Select":
                if self.fix_quant[i].get() == "N/A":
                    tkMessageBox.showwarning("Warning", "No quantity set for "+str(self.fix_com[i].get()))
                    self.fix_com[i].set("Select")
                else:
                    self.food2fix.append((self.fix_com[i].get(),self.fix_quant[i].get()))
                    print "Ration of "+str(self.fix_com[i].get())+" set to " +str(self.fix_quant[i].get())+ " gram/ration."
                    countfix+=1
        print " "
        self.allowshortfalls.set(1)
        self.sensible.set(0)

    def set_mod(self):
        '''
        User-added constraint: Set the transfer modality ratio
        '''

        f = self.cv_mod_fdp.get()
        mn = self.cv_mod_min.get()
        mx = self.cv_mod_max.get()
        if f == "Select":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            indices = self.listboxCV.curselection()
            time = []
            for i in indices:
                time.append(self.listboxCV.get(i))
            if time == []:
                time = list(self.horizon)
            for t in time:
                self.user_modality[f,t] = [mn,mx]
            print "Transfer modality ratio for " + f + " set to (" + mn + ", " + mx + ") for t=" , time
        print " "

    def set_nut(self):
        '''
        User-added constraint: Set the nutritional shortfall
        '''

        n = self.nutchoice.get()
        q = self.maxnut.get()
        if n == "Select" or q == "N/A":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            self.user_add_nut[n]=q
            print "You have bounded the shortfall of nutrient " + n + " at " + q +"%"
            self.nutchoice.set("Select")
            self.maxnut.set("N/A")
        print " "

    def set_obj(self):
        '''
        User-added constraint: Set an objective value
        '''

        s = self.statistic.get()
        mn = self.minstat.get()
        mx = self.maxstat.get()
        r = self.statrange.get()
        if s=="Select":
            tkMessageBox.showwarning("Error", "Not enough inputs")
        else:
            time = []
            if r=="Selected months":
                indices = self.listboxO.curselection()
                for i in indices:
                    time.append(self.listboxO.get(i))
                if time == []:
                    time = list(self.horizon)
                if s == "Lead Time (Avg)":
                    print "<<Warning>> Lead Time statistics are tricky to interpret when using individual periods; we recommend to use the Total or Average statistic"
            else:
                time.append(r)
            for t in time:
                self.mingoal[s,t] = mn
                self.maxgoal[s,t] = mx
            print "Added statistic constraint (stat, min, max, periods)"
            print (self.statistic.get(),mn,mx,time)
            self.statistic.set("Select")
            self.minstat.set("N/A")
            self.maxstat.set("N/A")
            self.statrange.set("Selected months")
        print " "

    def quick_load(self):
        '''
        Load temporarily saved user inputs (from quick_save(self))
        '''

        try:
            self.user_add_com = self.user_add_comCopy
            self.user_add_proc_int =self.user_add_proc_intCopy
            self.user_add_proc_loc =self.user_add_proc_locCopy
            self.user_add_route=self.user_add_routeCopy
            self.user_ex_com =self.user_ex_comCopy
            self.user_ex_proc_int=self.user_ex_proc_intCopy
            self.user_ex_proc_loc=self.user_ex_proc_locCopy
            self.user_ex_route=self.user_ex_routeCopy
            self.food2fix = self.food2fixCopy
            self.user_add_nut = self.user_add_nutCopy
            for k in self.fix_comCopy.items():
                self.fix_com[k[0]].set(k[1])
            for k in self.fix_quantCopy.items():
                self.fix_quant[k[0]].set(k[1])
            self.user_add_fg = self.user_add_fgCopy
            self.user_ex_fg = self.user_ex_fgCopy
            self.user_add_mincom = self.user_add_mincomCopy
            self.user_add_maxcom = self.user_add_maxcomCopy
            self.user_nut_minprot = self.user_nut_minprotCopy
            self.user_nut_maxprot = self.user_nut_maxprotCopy
            self.user_nut_minfat = self.user_nut_minfatCopy
            self.user_nut_maxfat = self.user_nut_maxfatCopy
            self.allowshortfalls.set(self.allowshortfallsCopy.get())
            self.gmo.set(self.gmoCopy.get())
            self.sensible.set(self.sensibleCopy.get())
            self.useforecasts.set(self.useforecastsCopy.get())
            self.supply_tact.set(self.supply_tactCopy.get())
            self.varbasket.set(self.varbasketCopy.get())
            self.mingoal=self.mingoalCopy.copy()
            self.maxgoal=self.maxgoalCopy.copy()
            self.user_ex_cv = self.user_ex_cvCopy
            self.user_add_cv = self.user_add_cvCopy
            self.user_cv_min.set(self.user_cv_minCopy.get())
            self.user_cv_max.set(self.user_cv_maxCopy.get())
            self.user_int_min.set(self.user_int_minCopy.get())
            self.user_int_max.set(self.user_int_maxCopy.get())
            self.user_reg_min.set(self.user_reg_minCopy.get())
            self.user_reg_max.set(self.user_reg_maxCopy.get())
            self.user_loc_min.set(self.user_loc_minCopy.get())
            self.user_loc_max.set(self.user_loc_maxCopy.get())
            self.user_cap_aloc = self.user_cap_alocCopy.copy()
            self.user_cap_util = self.user_cap_utilCopy.copy()
            self.user_add_ik = self.user_add_ikCopy.copy()
            for key in self.exp_pattern.keys():
                self.exp_pattern[key].set(self.exp_patternCopy[key].get())

            print "Constraints loaded from previously saved state"
            print " "
        except:
            print "<<ERROR>> Could not load"
            print "Have you used Quick Save yet?"
            print " "

    def quick_save(self):
        '''
        Temporarily save current user inputs (to be loaded using quick_load(self))
        Saved user inputs are lost when the app is closed
        '''

        self.user_add_comCopy = self.user_add_com.copy()
        self.user_add_proc_intCopy = self.user_add_proc_int.copy()
        self.user_add_proc_locCopy = self.user_add_proc_loc.copy()
        self.user_add_routeCopy = self.user_add_route.copy()
        self.user_ex_comCopy = list(self.user_ex_com)
        self.user_ex_proc_intCopy = self.user_ex_proc_int.copy()
        self.user_ex_proc_locCopy = self.user_ex_proc_loc.copy()
        self.user_ex_routeCopy = self.user_ex_route.copy()
        self.food2fixCopy = list(self.food2fix)
        self.user_add_nutCopy = self.user_add_nut.copy()
        self.fix_comCopy = {}
        self.fix_quantCopy = {}
        for k in self.fix_com.items():
            self.fix_comCopy[k[0]]=k[1].get()
        for k in self.fix_quant.items():
            self.fix_quantCopy[k[0]]=k[1].get()
        self.user_add_fgCopy = self.user_add_fg
        self.user_ex_fgCopy = list(self.user_ex_fg)
        self.user_add_mincomCopy = self.user_add_mincom
        self.user_add_maxcomCopy = self.user_add_maxcom
        self.user_nut_minprotCopy = self.user_nut_minprot
        self.user_nut_maxprotCopy = self.user_nut_maxprot
        self.user_nut_minfatCopy = self.user_nut_minfat
        self.user_nut_maxfatCopy = self.user_nut_maxfat
        self.allowshortfallsCopy = IntVar()
        self.allowshortfallsCopy.set(self.allowshortfalls.get())
        self.gmoCopy = IntVar()
        self.gmoCopy.set(self.gmo.get())
        self.sensibleCopy = IntVar()
        self.sensibleCopy.set(self.sensible.get())
        self.useforecastsCopy = IntVar()
        self.useforecastsCopy.set(self.useforecasts.get())
        self.supply_tactCopy = IntVar()
        self.supply_tactCopy.set(self.supply_tact.get())
        self.varbasketCopy = StringVar()
        self.varbasketCopy.set(self.varbasket.get())
        self.mingoalCopy = self.mingoal.copy()
        self.maxgoalCopy = self.maxgoal.copy()
        self.user_ex_cvCopy = self.user_ex_cv.copy()
        self.user_add_cvCopy = self.user_add_cv.copy()
        self.user_cv_minCopy = StringVar()
        self.user_cv_minCopy.set(self.user_cv_min.get())
        self.user_cv_maxCopy = StringVar()
        self.user_cv_maxCopy.set(self.user_cv_max.get())
        self.user_int_minCopy = StringVar()
        self.user_int_minCopy.set(self.user_int_min.get())
        self.user_int_maxCopy = StringVar()
        self.user_int_maxCopy.set(self.user_int_max.get())
        self.user_reg_minCopy = StringVar()
        self.user_reg_minCopy.set(self.user_reg_min.get())
        self.user_reg_maxCopy = StringVar()
        self.user_reg_maxCopy.set(self.user_reg_max.get())
        self.user_loc_minCopy = StringVar()
        self.user_loc_minCopy.set(self.user_loc_min.get())
        self.user_loc_maxCopy = StringVar()
        self.user_loc_maxCopy.set(self.user_loc_max.get())
        self.user_cap_alocCopy = self.user_cap_aloc.copy()
        self.user_cap_utilCopy = self.user_cap_util.copy()
        self.user_add_ikCopy = self.user_add_ik.copy()
        self.exp_patternCopy = {}
        for key in self.exp_pattern.keys():
            self.exp_patternCopy[key] = StringVar()
            self.exp_patternCopy[key].set(self.exp_pattern[key].get())

        print "Constraints saved -- Click the 'Quick Load' button to return to the current state"
        print " "

    def csv_convert(self):
        '''
        Convert all saved files to newest format
        Use with caution
        '''

        script_dir = os.path.dirname(os.path.abspath(__file__))
        mypath = os.path.join(script_dir, 'saved')

        folders = []
        for (dirpath, dirnames, filenames) in os.walk(mypath):
            folders.extend(dirnames)
            break
        print "Backing up old save files..."
        t = time.strftime("%Y-%m-%d %H:%M:%S", time.gmtime())
        t = t.replace(':','.')
        bup = os.path.join(mypath, 'backup ' + t)
        try:
            os.makedirs(bup)
        except OSError:
            pass # folder already exists
        for f in folders:
            if f.startswith("backup") or f.startswith("Backup"):
                continue
            o = os.path.join(mypath, f)
            d = os.path.join(bup, os.path.basename(o))
            shutil.copytree(o, d)
        print "Converting files..."
        e=0
        for f in folders:
            if f.startswith("backup") or f.startswith("Backup"):
                continue
            try:
                self.csvnamel.set(f)
                self.csvnames.set(f)
                self.csv_load_old()
                self.csv_save()
            except:
                print "<<ERROR>> Could not convert ",f
                e+=1
        if e==0:
            print "Succesfully converted all saved files to the newest format"
        else:
            print "Encountered " + str(e) + " errors during conversion of saved files to the newest format"
        print " "

    def csv_save(self):
        '''
        Save user input permanently to specified .csv file
        '''

        name = self.csvnames.get()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'saved')
        dest_dir = os.path.join(dest_dir, name)
        try:
            os.makedirs(dest_dir) # create ..\saved\name\  subfolder
        except OSError:
            pass # folder already exists

        path = os.path.join(dest_dir, "user_add_com.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Commodity","Min","Max"])
        for i in self.user_add_com.items():
            if i[1] != [0,1000]: # [0,1000] is the default value anyway
                c.writerow([i[0],i[1][0],i[1][1]])
        out.close()

        path = os.path.join(dest_dir, "user_add_proc_int.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Country","Incoterm","NDP","Commodity","mt","[t]"])
        for i in self.user_add_proc_int.items():
            row=[]
            for j in i[0]:
                row.append(j)
            row.append(i[1][0])
            for j in i[1][1]:
                row.append(j)
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_add_proc_loc.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Country","Incoterm","NDP","Commodity","mt","[t]"])
        for i in self.user_add_proc_loc.items():
            row=[]
            for j in i[0]:
                row.append(j)
            row.append(i[1][0])
            for j in i[1][1]:
                row.append(j)
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_add_route.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Origin","Destination","Commodity","mt","[t]"])
        for i in self.user_add_route.items():
            row=[]
            for j in i[0]:
                row.append(j)
            row.append(i[1][0])
            for j in i[1][1]:
                row.append(j)
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_add_cv.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Market","Commodity","t","mt"])
        for i in self.user_add_cv.items():
            c.writerow([i[0][0],i[0][1],i[0][2],i[1]])
        out.close()

        path = os.path.join(dest_dir, "user_add_fg.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Food Group","Min","Max"])
        for i in self.user_add_fg.items():
            c.writerow([i[0],i[1][0],i[1][1]])
        out.close()

        path = os.path.join(dest_dir, "user_add_nut.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Nutrient","Shortfall (%)"])
        for i in self.user_add_nut.items():
            c.writerow([i[0],i[1]])
        out.close()

        path = os.path.join(dest_dir, "user_add_ik.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Metric","Measurement","Period","Value"])
        for i in self.user_add_ik.items():
            c.writerow([i[0][0],i[0][1],i[0][2],i[1]])
        out.close()

        path = os.path.join(dest_dir, "user_ex_com.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Commodity"])
        for i in self.user_ex_com:
            c.writerow([i])
        out.close()

        path = os.path.join(dest_dir, "user_ex_proc_int.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Country","NDP","Commodity","[t]"])
        for i in self.user_ex_proc_int.items():
            row=[]
            for j in i[0]:
                row.append(j)
            for j in i[1]:
                row.append(j)
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_ex_proc_loc.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Country","NDP","Commodity","[t]"])
        for i in self.user_ex_proc_loc.items():
            row=[]
            for j in i[0]:
                row.append(j)
            for j in i[1]:
                row.append(j)
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_ex_route.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Origin","Destination","Commodity","[t]"])
        for i in self.user_ex_route.items():
            row=[]
            for j in i[0]:
                row.append(j)
            for j in i[1]:
                row.append(j)
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_ex_cv.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Market","Commodity","[t]"])
        for i in self.user_ex_cv.items():
            row=[i[0][0],i[0][1]]
            for j in i[1]:
                row.append(j)
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_ex_fg.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Food Group"])
        for i in self.user_ex_fg:
            c.writerow([i])
        out.close()

        path = os.path.join(dest_dir, "user_modality.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["FDP","t","Min","Max"])
        for i in self.user_modality.items():
            row = [i[0][0],i[0][1],i[1][0],i[1][1]]
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_cap_util.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Loc","t","Min","Max"])
        for i in self.user_cap_util.items():
            row = [i[0][0],i[0][1],i[1][0],i[1][1]]
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "user_cap_aloc.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Loc","t","Min","Max"])
        for i in self.user_cap_aloc.items():
            row = [i[0][0],i[0][1],i[1][0],i[1][1]]
            c.writerow(row)
        out.close()

        path = os.path.join(dest_dir, "food2fix.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Commodity","g/p/d"])
        for i in self.food2fix:
            c.writerow([i[0],i[1]])
        out.close()

        path = os.path.join(dest_dir, "fix_com.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Index","Commodity"])
        for i in self.fix_com.items():
            c.writerow([i[0],i[1].get()])
        out.close()

        path = os.path.join(dest_dir, "fix_quant.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Index","g/p/d"])
        for i in self.fix_quant.items():
            c.writerow([i[0],i[1].get()])
        out.close()

        path = os.path.join(dest_dir, "mingoal.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Statistic","t","Value"])
        for i in self.mingoal.items():
            c.writerow([i[0][0],i[0][1],i[1]])
        out.close()

        path = os.path.join(dest_dir, "maxgoal.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Statistic","t","Value"])
        for i in self.maxgoal.items():
            c.writerow([i[0][0],i[0][1],i[1]])
        out.close()

        path = os.path.join(dest_dir, "exp_pattern.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Group","Index","Value"])
        for i in self.exp_pattern.items():
            c.writerow([i[0][0],i[0][1],i[1].get()])
        out.close()

        path = os.path.join(dest_dir, "activities.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Activities"])
        for i in self.activities:
            c.writerow([i])
        out.close()

        path = os.path.join(dest_dir, "tactical_demand.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["Key","Check"])
        for i in self.tactboxes.items():
            c.writerow([i[0],i[1].get()])
        out.close()

        path = os.path.join(dest_dir, "single values.csv")
        out = open(path,"wb")
        c = csv.writer(out, dialect='excel')
        c.writerow(["tstart",self.tstart.get()])
        c.writerow(["tend",self.tend.get()])
        c.writerow(["ben",self.ben.get()])
        c.writerow(["allowshortfalls",self.allowshortfalls.get()])
        c.writerow(["gmo",self.gmo.get()])
        c.writerow(["sensible",self.sensible.get()])
        c.writerow(["useforecasts",self.useforecasts.get()])
        c.writerow(["supply_tact",self.supply_tact.get()])
        c.writerow(["varbasket",self.varbasket.get()])
        c.writerow(["user_add_mincom",self.user_add_mincom.get()])
        c.writerow(["user_add_maxcom",self.user_add_maxcom.get()])
        c.writerow(["user_cv_min",self.user_cv_min.get()])
        c.writerow(["user_cv_max",self.user_cv_max.get()])
        c.writerow(["user_int_min",self.user_int_min.get()])
        c.writerow(["user_int_max",self.user_int_max.get()])
        c.writerow(["user_reg_min",self.user_reg_min.get()])
        c.writerow(["user_reg_max",self.user_reg_max.get()])
        c.writerow(["user_loc_min",self.user_loc_min.get()])
        c.writerow(["user_loc_max",self.user_loc_max.get()])
        c.writerow(["user_nut_minprot",self.user_nut_minprot.get()])
        c.writerow(["user_nut_maxprot",self.user_nut_maxprot.get()])
        c.writerow(["user_nut_minfat",self.user_nut_minfat.get()])
        c.writerow(["user_nut_maxfat",self.user_nut_maxfat.get()])
        out.close()

        print "User constraints saved to "+dest_dir
        print " "
        self.scenname.set(name)
        self.update_csv()  # update the load csv menu

    def csv_load(self):
        '''
        Load user input from specified .csv file
        '''

        name = self.csvnamel.get()
        self.reset()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'saved')
        dest_dir = os.path.join(dest_dir, name)
        e = 0

        try:
            path = os.path.join(dest_dir, "single values.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            self.tstart.set(c.next()[1])
            self.tend.set(c.next()[1])
            self.ben.set(c.next()[1])
            self.allowshortfalls.set(c.next()[1])
            self.gmo.set(c.next()[1])
            self.sensible.set(c.next()[1])
            self.useforecasts.set(c.next()[1])
            self.supply_tact.set(c.next()[1])
            self.varbasket.set(c.next()[1])
            self.user_add_mincom.set(c.next()[1])
            self.user_add_maxcom.set(c.next()[1])
            self.user_cv_min.set(c.next()[1])
            self.user_cv_max.set(c.next()[1])
            self.user_int_min.set(c.next()[1])
            self.user_int_max.set(c.next()[1])
            self.user_reg_min.set(c.next()[1])
            self.user_reg_max.set(c.next()[1])
            self.user_loc_min.set(c.next()[1])
            self.user_loc_max.set(c.next()[1])
            self.user_nut_minprot.set(c.next()[1])
            self.user_nut_maxprot.set(c.next()[1])
            self.user_nut_minfat.set(c.next()[1])
            self.user_nut_maxfat.set(c.next()[1])
            out.close()
        except:
            print "<<<Error>>> Could not load 'single values.csv'"
            e += 1

        try:
            path = os.path.join(dest_dir, "activities.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            self.activities = []
            for i in c:
                self.activities.append(i[0])
            out.close()
        except:
            print "<<<Error>>> Could not load activities.csv"
            e += 1
        self.listbox_act.delete(0,END)
        for b in self.beneficiaries:
            if b == self.ben.get():
                continue
            self.listbox_act.insert(END,b)
        for i in range(self.listbox_act.size()):
            if self.listbox_act.get(i) in self.activities:
                self.listbox_act.itemconfig(i, background="light sky blue")
            else:
                self.listbox_act.itemconfig(i, background="white")
        self.act_button.configure(text="Select (" + str(len(self.activities)) + ")")

        try:
            path = os.path.join(dest_dir, "user_add_com.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_com[i[0]]=[i[1],i[2]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_com.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_proc_int.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_proc_int[i[0],i[1],i[2],i[3]]=[i[4],i[5:]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_proc_int.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_proc_loc.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_proc_loc[i[0],i[1],i[2],i[3]]=[i[4],i[5:]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_proc_loc.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_route.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_route[i[0],i[1],i[2]]=[i[3],i[4:]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_route.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_cv.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_cv[i[0],i[1],i[2]]=i[3]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_cv.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_fg.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_fg[i[0]]=[i[1],i[2]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_fg.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_nut.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_nut[i[0]]=i[1]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_nut.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_ik.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_ik[i[0],i[1],i[2]]=i[3]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_ik.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_com.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_com.append(i[0])
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_com.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_proc_int.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_proc_int[i[0],i[1],i[2]]=i[3:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_proc_int.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_proc_loc.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_proc_loc[i[0],i[1],i[2]]=i[3:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_proc_loc.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_route.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_route[i[0],i[1],i[2]]=i[3:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_route.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_cv.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_cv[i[0],i[1]]=i[2:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_cv.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_fg.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_fg.append(i[0])
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_fg.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_modality.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_modality[i[0],i[1]]=[i[2],i[3]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_modality.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_cap_util.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_cap_util[i[0],i[1]]=[i[2],i[3]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_cap_util.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_cap_aloc.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_cap_aloc[i[0],i[1]]=[i[2],i[3]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_cap_aloc.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "food2fix.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.food2fix.append([i[0],i[1]])
            out.close()
        except:
            print "<<<Error>>> Could not load food2fix.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "fix_com.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.fix_com[int(i[0])].set(i[1])
            out.close()
        except:
            print "<<<Error>>> Could not load fix_com.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "fix_quant.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.fix_quant[int(i[0])].set(i[1])
            out.close()
        except:
            print "<<<Error>>> Could not load fix_quant.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "mingoal.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.mingoal[i[0],i[1]]=i[2]
            out.close()
        except:
            print "<<<Error>>> Could not load mingoal.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "maxgoal.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.maxgoal[i[0],i[1]]=i[2]
            out.close()
        except:
            print "<<<Error>>> Could not load maxgoal.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "exp_pattern.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.exp_pattern[i[0],int(i[1])].set(i[2])
            out.close()
        except:
            print "<<<Error>>> Could not load exp_pattern.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "tactical_demand.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.tactboxes[i[0]].set(i[1])
            out.close()
        except:
            print "<<<Error>>> Could not load tactical_demand.csv"
            e += 1

        self.scenname.set(name)
        self.csvnames.set(name)

        if e==0:
            print "Succesfully loaded " + name
        else:
            print str(e) + " errors encountered during loading of " + name
        print " "

    def csv_load_old(self):
        '''
        Load user input from .csv file using previous methodology (for conversion)
        '''

        name = self.csvnamel.get()
        self.reset()
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'saved')
        dest_dir = os.path.join(dest_dir, name)
        e = 0

        try:
            path = os.path.join(dest_dir, "single values.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            self.tstart.set(c.next()[1])
            self.tend.set(c.next()[1])
            self.ben.set(c.next()[1])
            self.allowshortfalls.set(c.next()[1])
            self.gmo.set(c.next()[1])
            self.sensible.set(c.next()[1])
            self.useforecasts.set(c.next()[1])
            self.varbasket.set(c.next()[1])
            self.user_add_mincom.set(c.next()[1])
            self.user_add_maxcom.set(c.next()[1])
            self.user_cv_min.set(c.next()[1])
            self.user_cv_max.set(c.next()[1])
            self.user_int_min.set(c.next()[1])
            self.user_int_max.set(c.next()[1])
            self.user_reg_min.set(c.next()[1])
            self.user_reg_max.set(c.next()[1])
            self.user_loc_min.set(c.next()[1])
            self.user_loc_max.set(c.next()[1])
            self.obj2.set(c.next()[1])
            self.objmin.set(c.next()[1])
            self.objmax.set(c.next()[1])
            self.increment.set(c.next()[1])
            self.user_nut_minprot.set(c.next()[1])
            self.user_nut_maxprot.set(c.next()[1])
            self.user_nut_minfat.set(c.next()[1])
            self.user_nut_maxfat.set(c.next()[1])
            out.close()
        except:
            print "<<<Error>>> Could not load 'single values.csv'"
            e += 1

        try:
            path = os.path.join(dest_dir, "activities.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            self.activities = []
            for i in c:
                self.activities.append(i[0])
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_com.csv"
            e += 1
        self.listbox_act.delete(0,END)
        for b in self.beneficiaries:
            if b == self.ben.get():
                continue
            self.listbox_act.insert(END,b)
        for i in range(self.listbox_act.size()):
            if self.listbox_act.get(i) in self.activities:
                self.listbox_act.itemconfig(i, background="light sky blue")
            else:
                self.listbox_act.itemconfig(i, background="white")
        self.act_button.configure(text="Select (" + str(len(self.activities)) + ")")

        try:
            path = os.path.join(dest_dir, "user_add_com.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_com[i[0]]=[i[1],i[2]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_com.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_proc_int.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_proc_int[i[0],i[1],i[2],i[3]]=[i[4],i[5:]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_proc_int.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_proc_loc.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_proc_loc[i[0],i[1],i[2],i[3]]=[i[4],i[5:]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_proc_loc.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_route.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_route[i[0],i[1],i[2]]=[i[3],i[4:]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_route.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_cv.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_cv[i[0],i[1],i[2]]=i[3]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_cv.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_fg.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_fg[i[0]]=[i[1],i[2]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_fg.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_nut.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_nut[i[0]]=i[1]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_nut.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_add_ik.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_add_ik[i[0],i[1],i[2]]=i[3]
            out.close()
        except:
            print "<<<Error>>> Could not load user_add_ik.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_com.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_com.append(i[0])
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_com.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_proc_int.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_proc_int[i[0],i[1],i[2]]=i[3:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_proc_int.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_proc_loc.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_proc_loc[i[0],i[1],i[2]]=i[3:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_proc_loc.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_route.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_route[i[0],i[1],i[2]]=i[3:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_route.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_cv.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_cv[i[0],i[1]]=i[2:]
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_cv.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_ex_fg.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_ex_fg.append(i)
            out.close()
        except:
            print "<<<Error>>> Could not load user_ex_fg.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_modality.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_modality[i[0],i[1]]=[i[2],i[3]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_modality.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_cap_util.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_cap_util[i[0],i[1]]=[i[2],i[3]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_cap_util.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "user_cap_aloc.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.user_cap_aloc[i[0],i[1]]=[i[2],i[3]]
            out.close()
        except:
            print "<<<Error>>> Could not load user_cap_aloc.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "food2fix.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.food2fix.append([i[0],i[1]])
            out.close()
        except:
            print "<<<Error>>> Could not load food2fix.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "fix_com.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.fix_com[int(i[0])].set(i[1])
            out.close()
        except:
            print "<<<Error>>> Could not load fix_com.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "fix_quant.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.fix_quant[int(i[0])].set(i[1])
            out.close()
        except:
            print "<<<Error>>> Could not load fix_quant.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "mingoal.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.mingoal[i[0],i[1]]=i[2]
            out.close()
        except:
            print "<<<Error>>> Could not load mingoal.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "maxgoal.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.maxgoal[i[0],i[1]]=i[2]
            out.close()
        except:
            print "<<<Error>>> Could not load maxgoal.csv"
            e += 1

        try:
            path = os.path.join(dest_dir, "exp_pattern.csv")
            out = open(path,"rb")
            c = csv.reader(out, dialect='excel')
            next(c,None)
            for i in c:
                self.exp_pattern[i[0],int(i[1])].set(i[2])
            out.close()
        except:
            print "<<<Error>>> Could not load exp_pattern.csv"
            e += 1

        self.scenname.set(name)
        self.csvnames.set(name)

        if e==0:
            print "Succesfully loaded " + name
        else:
            print str(e) + " errors encountered during loading of " + name
        print " "

    def csv_outputs(self,LOC,NAME):
        '''
        Writes all model outputs and results of a single solution to .csv files:
        > Food Basket
        > Routing
        > Sourcing
        > Statistics
        '''

        NAME = NAME.replace('/','-') # /-signs create subfolders (occurs f.i. with "SORGHUM/MILLET")
        NAME = NAME.replace('<','leq')
        NAME = NAME.replace('>','geq')
        LOC = LOC.replace('/','-') # /-signs create subfolders (occurs f.i. with "SORGHUM/MILLET")
        LOC = LOC.replace('<','leq')
        LOC = LOC.replace('>','geq')
        print "Writing output files..."
        try:
            os.makedirs(LOC) # create output folder
        except OSError:
            pass # folder already exists
        path = os.path.join(LOC, "") # \output\scenname\
        script_dir = os.path.dirname(os.path.abspath(__file__))
        macro_loc = os.path.join(script_dir, 'data')
        macro_loc = os.path.join(macro_loc, 'Scenario Dashboard.xlsm')
        try:
            shutil.copyfile(macro_loc,path+"Generate Output (" + NAME + ").xlsm") # Copy the output macro to the folder
        except IOError:
            print "The 'Generate Output' file is currently opened, please reload its data manually!"

        path_fb = open(path+"Food Basket.csv",'wb')
        path_src = open(path+"Sourcing.csv",'wb')
        path_route = open(path+"Routing.csv", 'wb')
        path_stat = open(path+"Statistics.csv", 'wb')

        # Write routing file
        csv_file = csv.writer(path_route, dialect='excel')
        csv_file.writerow(["O-Type","Origin","D-Type","Destination","Commodity","Time","Metric Tonnes","Cost","Lead Time"])
        for arc in self.arcs.items():
            if(value(self.F[arc[0]]) != 0 and arc[0][1] in (self.DPs+self.EDPs+self.FDPs) and arc[0][0] not in self.sources): # i.e. all non-procurement movements that have a value
                if( value(self.F[arc[0]]) == None ): # Weird exception, under investigation. Popped up for an RUTF commodity once
                    print "<<<ERROR>>> Something weird's going on with "+arc[0][2]+" from "+arc[0][0] + " to " + arc[0][1]
                    self.errors+=1
                else:
                    csv_file.writerow([self.type[arc[0][0]], arc[0][0], self.type[arc[0][1]], arc[0][1], arc[0][2], self.horizon[arc[0][3]], str(value(self.F[arc[0]])),str(value(self.F[arc[0]])*arc[1]),str(self.dur[arc[0][0],arc[0][1],arc[0][2]])])

        # Write sourcing file
        ltm = 0
        csv_file = csv.writer(path_src, dialect='excel')
        csv_file.writerow(["Origin Country","Delivery Terms","Incoterm","Incoterm Country","Named Delivery Place","Commodity","Time","Metric Tonnes","$/mt","Total Cost","Lead Time"])
        for arc in self.arcs.items():
            if(value(self.F[arc[0]]) != 0 and arc[0][0] in self.sources): # i.e. all procurement movements that have a value
                oc = arc[0][0][:-6]
                inco = arc[0][0][-3:]
                ndp = arc[0][1]
                i = ndp.index("(") + 1
                delterms = inco + " " + ndp[:i-1]
                incocountry = ndp[i:-1]
                csv_file.writerow([oc, delterms, inco, incocountry, ndp, arc[0][2], self.horizon[arc[0][3]], str(value(self.F[arc[0]])), str(arc[1]), str(value(self.F[arc[0]])*arc[1]), str(self.quick[arc[0][0],arc[0][1]])])
                if self.quick[arc[0][0],arc[0][1]] > ltm:
                    ltm = self.quick[arc[0][0],arc[0][1]]

        # Write food basket file
        csv_file = csv.writer(path_fb, dialect='excel')
        self.usedcoms = [] # This subset makes it easier to show food basket changes across time
        for k in self.commodities:
            for t in self.hor:
                if(value(self.R[k][t]) > 0):
                    self.usedcoms.append(k)
        self.usedcoms = list(set(self.usedcoms))
        self.usedcoms.sort()
        csv_file.writerow(["Period","Type 1","Item","Type 2","Nutrient","Value"])
        for t in self.hor:
            val = {}
            p = self.horizon[t]
            if self.dem[self.ben.get(),p] == 0:
                continue
            for k in self.usedcoms:
                csv_file.writerow([p, "Commodity", k, "Macro", " g/p/d", value(self.R[k][t])])
                for l in self.nutrients:
                    if l in ("ENERGY (kcal)", "PROTEIN (g)", "FAT    (g)"):
                        type2 = "Macro"
                    else:
                        type2 = "Micro"
                    v = value(self.R[k][t]) * self.nutval[k,l]/100.0
                    csv_file.writerow([p, "Commodity", k, type2, l, v])
                    if l in val.keys():
                        val[l] += v
                    else:
                        val[l] = v
            if len(self.usedcoms) == 0: # empty basket
                for l in self.nutrients:
                    val[l] = 0
            for l in self.nutrients:
                if l in ("ENERGY (kcal)", "PROTEIN (g)", "FAT    (g)"):
                    type2 = "Macro"
                else:
                    type2 = "Micro"
                csv_file.writerow([p, "Statistic", "Nutritional Content", type2, l, val[l]])
                csv_file.writerow([p, "Statistic", "Nutritional Requirement", type2, l, self.nutreq[self.ben.get(),l]])
                csv_file.writerow([p, "Statistic", "Nutritional Shortfall", type2, l, value(self.S[l][t])])
                csv_file.writerow([p, "Statistic", "NVS Contribution", type2, l, (1-value(self.S[l][t]))])
                csv_file.writerow([p, "Statistic", "Percentage Supplied", type2, l, val[l]/self.nutreq[self.ben.get(),l]])

        # Write statistics file
        csv_file = csv.writer(path_stat, dialect='excel')
        header = ["Statistic","Total","Average"]
        for t in self.horizon:
            header.append(t)
        csv_file.writerow(header)
        for s in sorted(self.stats.items()): # s = ("name")(statistic)
            row = []
            row.append(s[0])
            if s[0] != "Lead Time":
                row.append(str(value(s[1]["Total"])))
                row.append(str(value(s[1]["Average"])))
                for t in self.hor:
                    row.append(str(value(s[1][t])))
            else:
                row.append(str(ltm))
                if value(self.MT["Total"]) > 0:
                    row.append(str(value(self.LTsum["Total"])/value(self.MT["Total"])))
                else:
                    row.append("")
                for t in self.hor:
                    ltm = 0
                    for arc in self.cost.keys():
                        if value(self.F[arc[0],arc[1],arc[2],t]) > 0 and arc[0] in self.sources:
                            if self.quick[arc[0],arc[1]] > ltm:
                                ltm = self.quick[arc[0],arc[1]]
                    row.append(str(ltm))
            csv_file.writerow(row)
        csv_file.writerow(" ")

        header = ["Location","Statistic"]
        temp = {}
        for t in self.horizon:
            header.append(t)
        csv_file.writerow(header)
        for i in self.DPs:
            row1 = [i,"Throughput"]
            row2 = ["", "Capacity"]
            for t in self.hor:
                mt_t = value(self.LOAD[i,t])
                mt_c = self.nodecap[i,self.horizon[t]]
                row1.append(str(mt_t))
                row2.append(str(mt_c))
                if t in temp.keys():
                    temp[t][0] += mt_t
                    temp[t][1] += mt_c
                else:
                    temp[t] = [mt_t, mt_c]
            csv_file.writerow(row1)
            csv_file.writerow(row2)
        row = ["All DPs", "Utilisation (%)"]
        for t in self.hor:
            row.append(str(temp[t][0]/temp[t][1]))
        csv_file.writerow(row)
        temp = {}
        for i in self.EDPs:
            row1 = [i,"Throughput"]
            row2 = ["", "Capacity"]
            for t in self.hor:
                mt_t = value(self.LOAD[i,t])
                mt_c = self.nodecap[i,self.horizon[t]]
                row1.append(str(mt_t))
                row2.append(str(mt_c))
                if t in temp.keys():
                    temp[t][0] += mt_t
                    temp[t][1] += mt_c
                else:
                    temp[t] = [mt_t, mt_c]
            csv_file.writerow(row1)
            csv_file.writerow(row2)
        row = ["All EDPs", "Utilisation (%)"]
        for t in self.hor:
            row.append(str(temp[t][0]/temp[t][1]))
        csv_file.writerow(row)
        for i in self.FDPs:
            row1 = [i, "Food (mt)"]
            row2 = ["", "C&V (mt)"]
            row3 = ["", "C&V (%)"]
            for t in self.hor:
                mt_f = value(self.LOAD_F[i,t])
                mt_cv = value(self.LOAD_CV[i,t])
                row1.append(str(mt_f))
                row2.append(str(mt_cv))
                if mt_f + mt_cv > 0:
                    row3.append(str(mt_cv/(mt_f+mt_cv)))
                else:
                    row3.append("")
            csv_file.writerow(row1)
            csv_file.writerow(row2)
            csv_file.writerow(row3)

        path_fb.close()
        path_src.close()
        path_route.close()
        path_stat.close()
        print "Output files saved to: " + LOC
        print " "

    def csv_benchmarks(self,LOC,NAME):
        '''
        Writes cross-comparison statistics (KPIs) to a .csv file for the scenarios in the current analysis
        '''

        try:
            os.makedirs(LOC) # create output folder
        except OSError:
            pass # folder already exists
        script_dir = os.path.dirname(os.path.abspath(__file__))
        macro_loc = os.path.join(script_dir, 'data')
        macro_loc = os.path.join(macro_loc, 'Benchmarks Dashboard.xlsm')
        shutil.copyfile(macro_loc,os.path.join(LOC,"Generate Output (" + NAME + ").xlsm")) # Copy the output macro to the folder
        path = os.path.join(LOC,"Benchmarks.csv")

        try:
            out = open(path, "wb")
            c = csv.writer(out, dialect='excel')
            header=["Scenario: "]
            for item in self.solutions.items(): # item = (scenname) ([stat, val])
                for key in item[1]:
                    header.append(key[0][:-2])
                break
            c.writerow(header)
            for i in sorted(self.solutions.items()):
                row=[]
                row.append(i[0])
                for col in i[1]: # col = [statistic, value]
                    if str(col[1])[0]=="$":
                        row.append(col[1][1:]) # csv doesn't handle the $ very well
                    else:
                        row.append(col[1])
                c.writerow(row)
            out.close()
            print "Benchmarks saved to "+str(path)
            print " "
        except IOError:
            print "File is currently opened, please close it ;-)"

    def display_solution(self,VAR):
        '''
        Display variable values for last scenario
        '''

        indices = self.listboxSOL.curselection()
        time = []
        for i in indices:
            time.append(self.listboxCV.get(i))
        print "Showing " + VAR + " for periods: ", time

        if VAR == "Food Basket":
            for period in time:
                print "Month:",period
                t = self.horizon.index(period)
                for k in self.commodities:
                    if value(self.R[k][t]) > 0 :
                        print "%-35s %s" % (k, self.fmt_wcommas(value(self.R[k][t]))[1:].rjust(6))

        if VAR == "Nutrient Shortfalls":
            for period in time:
                print "Month:",period
                t = self.horizon.index(period)
                for l in self.nutrients:
                    if value(self.S[l][t]) > 0 :
                        print "%-20s %s" % (l, '{:.2%}'.format(value(self.S[l][t])).rjust(6))

        if VAR == "Sourcing Strategy":
            for period in time:
                print "Month:",period
                t = self.horizon.index(period)
                for arc in self.proccap:
                    if value(self.F[arc[0],arc[1],arc[2],t]) > 0 :
                        print arc
                        print self.fmt_wcommas(value(self.F[arc[0],arc[1],arc[2],t]))[1:]

        print " "

    def display_outputs(self,NAME):
        '''
        Display key outputs for the optimised scenario
        '''

        self.solutions[NAME]=[]
        print "------------------------------------------"
        print "--------------- Benchmarks ---------------"
        print "------------------------------------------"
        ben = []
        for t in self.hor:
            b = sum(self.dem[b,self.horizon[t]] for b in (self.activities+[self.ben.get()]))*self.scaleup
            if b > 0:
                ben.append(b)
        tc = self.fmt_wcommas(value(self.TC["Total"]/float(len(ben))))
        self.solutions[NAME].append(["Monthly Costs (Avg): ",tc])
        cpbpm = sum(ben)/float(len(ben)) # average beneficiaries per month
        cpbpm = self.fmt_wcommas(value(self.TC["Total"])/cpbpm/len(ben))
        self.solutions[NAME].append(["USD/ben/month (Avg): ", cpbpm])
        nv = self.fmt_wcommas(value(self.NVS["Total"])/float(len(self.hor)-len(self.empty)))[1:]
        self.solutions[NAME].append(["NVS (Avg): ",nv])
        perc = '{:.2%}'.format(value(self.NVS["Total"])/float(len(self.hor)-len(self.empty))/11)
        self.solutions[NAME].append(["Nutrients supplied (%): ",perc])
        kcal = self.fmt_wcommas(value(self.KCAL["Average"]))[1:]
        self.solutions[NAME].append(["Kcal (Avg): ",kcal])
        if value(self.KCAL["Average"]) > 0:
            ep = '{:.2%}'.format(value(self.PROT["Average"]) / value(self.KCAL["Average"]))
            ef = '{:.2%}'.format(value(self.FAT["Average"]) / value(self.KCAL["Average"]))
        else:
            ep = '{:.2%}'.format(0)
            ef = '{:.2%}'.format(0)
        self.solutions[NAME].append(["Energy (Protein): ",ep])
        self.solutions[NAME].append(["Energy (Fat): ",ef])
        if value(self.MT["Total"]) != 0:
            cvp = '{:.2%}'.format(value(self.MT_CV["Total"])/value(self.MT["Total"]))
            lpp = '{:.2%}'.format(value(self.MT_L["Total"])/value(self.MT["Total"]))
            alt = self.fmt_wcommas(value(self.LTsum["Total"])/value(self.MT["Total"]))[1:]
        else:
            cvp = "0.00%"
            lpp = "0.00%"
            alt = "0"
        self.solutions[NAME].append(["% C&V Proc (Avg): ",cvp])
        self.solutions[NAME].append(["% Loc Proc (Avg): ",lpp])
        self.solutions[NAME].append(["Lead Time (Avg): ",alt])
        dp_t = sum(value(self.LOAD[i,t]) for i in self.DPs for t in self.hor)
        dp_c = sum(self.nodecap[i,t] for i in self.DPs for t in self.horizon)
        dp_u = '{:.2%}'.format(dp_t/dp_c)
        self.solutions[NAME].append(["DP Utilisation (%): ",dp_u])
        edp_t = sum(value(self.LOAD[i,t]) for i in self.EDPs for t in self.hor)
        edp_c = sum(self.nodecap[i,t] for i in self.EDPs for t in self.horizon)
        edp_u = '{:.2%}'.format(edp_t/edp_c)
        self.solutions[NAME].append(["EDP Utilisation (%): ",edp_u])
        try:
            ik_usd = self.fmt_wcommas(value(self.TC_IK["Total"]))
            ik_mt = self.fmt_wcommas(value(self.MT_IK["Total"]))[1:]
        except:
            ik_usd = "$0.00"
            ik_mt = "0"
        self.solutions[NAME].append(["US IK ($): ",ik_usd])

        self.solutions[NAME].append(["US IK (mt): ",ik_mt])

        max1 = 0
        max2 = 0
        for item in self.solutions[NAME]:
            if len(item[0]) > max1:
                max1 = len(item[0])
            if len(item[1]) > max2:
                max2 = len(item[1])
        for item in self.solutions[NAME]:
            print "%-25s %s" % (item[0].ljust(max1), item[1].rjust(max2))

        print "------------------------------------------"
        if self.errors != 0:
            print "<<WARNING>> NOT ALL USER CONSTRAINTS WERE ACCEPTED <<WARNING>>"
            print "Scroll through the output to find out where it went wrong (look for <<ERROR>>)"
        print " "

        self.disp[NAME]=list(self.solutions[NAME]) # The subset of self.solutions that is displayed in the GUI
        pc = self.fmt_wcommas(value(self.PC["Total"]))
        self.solutions[NAME].append(["Procurement Costs: ",pc])
        trc = self.fmt_wcommas(value(self.TR["Total"]))
        self.solutions[NAME].append(["Transportation Costs: ",trc])
        hand = self.fmt_wcommas(value(self.HC["Total"]))
        self.solutions[NAME].append(["Handling Costs: ",hand])
        od = self.fmt_wcommas(value(self.ODOC["Total"]))
        self.solutions[NAME].append(["ODOC Costs: ",od])
        nvmin = self.fmt_wcommas(min(value(self.NVS[t]) for t in self.hor if t not in self.empty))[1:]
        self.solutions[NAME].append(["NVS (Min): ",nvmin])
        nvmax = self.fmt_wcommas(max(value(self.NVS[t]) for t in self.hor if t not in self.empty))[1:]
        self.solutions[NAME].append(["NVS (Max): ",nvmax])
        ltm = 0
        for arc in self.arcs.keys():
            if value(self.F[arc]) > 0 and arc[0] in self.sources:
                if self.quick[arc[0],arc[1]] > ltm:
                    ltm = self.quick[arc[0],arc[1]]
        self.solutions[NAME].append(["Lead Time (Max): ",int(ltm)])
        for k in self.commodities:
            self.solutions[NAME].append([k + ": ", sum(value(self.R[k][t]) for t in self.hor)/(len(self.hor)-len(self.empty))]) # add average ration size for each k

    def display_benchmarks(self):
        '''
        Pops up a window with KPIs for the scenarios in the current analysis
        '''

        try:
            self.close(self.benchwin) # it may still be open from the last analysis
        except:
            None
        self.benchwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.benchwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.benchwin))
        self.benchwin.columnconfigure(0, weight=1)
        self.benchwin.rowconfigure(0, weight=1)

        header=["Scenario: "]
        for item in self.disp.items(): # item = (scenname) ([stat, val])
            for key in item[1]:
                header.append(key[0])
            break
        data = []
        for s in sorted(self.disp.items()):
            row = [s[0]]
            for i in s[1]:
                row.append(i[1])
            data.append(row)

        #h = min(20, len(self.disp.keys()))
        self.bmcanvas = Canvas(self.benchwin, width=800, height=200, background=self.bgcolor, bd=0, highlightthickness=0)
        self.bmframe = ttk.Frame(self.bmcanvas)
        self.bmframe.bind("<Configure>", self.OnFrameConfigure)
        self.bmtable = ttk.Treeview(self.bmframe, columns=header, show="headings", height=len(self.disp.keys()))
        self.bmtable.pack(fill="both", expand=True)
        self.bmcanvas.grid(row=0, column=0, sticky=W+E+N+S)
        self.bmcanvas.columnconfigure(0, weight=1)
        self.bmcanvas.rowconfigure(0, weight=1)
        self.bmcanvas.create_window((0,0), window=self.bmframe, anchor="nw", tags="self.bmframe")
        self.vsb = Scrollbar(self.benchwin, orient="vertical", command=self.bmcanvas.yview)
        self.hsb = Scrollbar(self.benchwin, orient="horizontal", command=self.bmcanvas.xview)
        self.bmcanvas.configure(xscrollcommand=self.hsb.set, yscrollcommand=self.vsb.set)
        self.vsb.grid(row=0, column=1, sticky="NS")
        self.hsb.grid(row=1, column=0, sticky="EW")

        self.bmtable.tag_configure('all', foreground='black')
        for col in header:
            self.bmtable.heading(col, text=col, command=lambda c=col: self.sortTree(self.bmtable, c, 0))
            self.bmtable.column(col, width=6*len(col))
        for item in data:
            self.bmtable.insert('', 'end', values=item, tags=('all',))
            for ix, val in enumerate(item):
                col_w = 6*len(val)
                if self.bmtable.column(header[ix],width=None) < col_w:
                    self.bmtable.column(header[ix], width=col_w) # Increase column width if necessary

        self.benchwin.deiconify()

    def display_inputs(self):
        '''
        Displays the current user inputs
        '''

        print "+---------------------+"
        print "| LIST OF USER INPUTS |"
        print "+---------------------+"
        print "Activity to be optimised: " , self.ben.get()
        if len(self.activities) > 0:
            print "Additional activities to be supplied:"
            for b in self.activities:
                print "> " + b
        if self.allowshortfalls.get() != 0:
            print "Shortfalls allowed: " , "Yes"
        if self.sensible.get() != 1:
            print "Sensible food basket: " , "No"
        if self.useforecasts.get() != 1:
            print "Use forecasts: " , "No"
        if self.varbasket.get() != "Variable":
            print "Basket variability: " , self.varbasket.get()
        if self.gmo.get() != 1:
            print "GMO allowed: " , "No"
        if len(self.user_add_proc_int.items()) > 0:
            print "Included Procurement Decisions (int/reg):"
            for i in sorted(self.user_add_proc_int.items()):
                print i
        if len(self.user_add_proc_loc.items()) > 0:
            print "Included Procurement Decisions (loc):"
            for i in sorted(self.user_add_proc_loc.items()):
                print i
        if len(self.user_add_cv.items()) > 0:
            print "Included Procurement Decisions (C&V):"
            for i in sorted(self.user_add_cv.items()):
                print i
        if len(self.user_add_route.items()) > 0:
            print "Included Routing Decisions:"
            for i in sorted(self.user_add_route.items()):
                print i
        if len(self.user_cap_util.items()) > 0:
            print "Included Capacity Utilisation Decisions:"
            for i in sorted(self.user_cap_util.items()):
                print i
        if len(self.user_cap_aloc.items()) > 0:
            print "Included Capacity Allocation Decisions:"
            for i in sorted(self.user_cap_aloc.items()):
                print i
        if len(self.user_add_fg.items()) > 0:
            print "Included Food Groups:"
            for g in sorted(self.user_add_fg.items()):
                print g
        if len(self.user_add_ik.items()) > 0:
            print "Included In-Kind Donations:"
            for i in sorted(self.user_add_ik.items()):
                print i
        if len(self.user_ex_com) > 0:
            print "Excluded commodities:"
            for i in sorted(self.user_ex_com):
                print i
        if len(self.user_ex_proc_int.items()) > 0:
            print "Excluded Procurement Decisions (int):"
            for i in sorted(self.user_ex_proc_int.items()):
                print i
        if len(self.user_ex_proc_loc.items()) > 0:
            print "Excluded Procurement Decisions (loc):"
            for i in sorted(self.user_ex_proc_loc.items()):
                print i
        if len(self.user_ex_cv.items()) > 0:
            print "Excluded Procurement Decisions (C&V):"
            for i in sorted(self.user_ex_cv.items()):
                print i
        i = [self.user_int_min.get(),self.user_int_max.get()]
        r = [self.user_reg_min.get(),self.user_reg_max.get()]
        l = [self.user_loc_min.get(),self.user_loc_max.get()]
        c = [self.user_cv_min.get(),self.user_cv_max.get()]
        if (i!=["0","100"] or r!=["0","100"] or l!=["0","100"] or c!=["0","100"]):
            lst = []
            lst.append(["International Procurement",i])
            lst.append(["Regional Procurement",r])
            lst.append(["Local Procurement",l])
            lst.append(["C&V Procurement",c])
            print "Procurement Allocations:"
            for item in lst:
                print "%-25s %s" % (item[0], item[1])
        if len(self.user_ex_route.items()) > 0:
            print "Excluded Routes:"
            for i in sorted(self.user_ex_route.items()):
                print i
        if len(self.user_ex_fg) > 0:
            print "Excluded Food Groups:"
            for g in sorted(self.user_ex_fg):
                print g
        if len(self.food2fix) > 0:
            print "Fixed Ration Sizes:"
            for i in sorted(self.food2fix):
                print i
        for i in self.user_add_com.items():
            if i[1] != ["0","1000"] and i[1] != [0,1000]:
                print "Adjusted Commodity Boundaries:"
                for j in sorted(self.user_add_com.items()):
                    if j[1] != ["0","1000"] and j[1] != [0,1000] and j[1] != ["N/A","N/A"]:
                        print j
                break
        if len(self.user_add_nut.items()) > 0:
            print "Bounds on Nutritional Shortfalls:"
            for i in sorted(self.user_add_nut.items()):
                if i[1] != "N/A":
                    print i
        if len(self.mingoal.keys()) > 0:
            print "Bounds on Statistics:"
            for o in sorted(self.mingoal.keys()):
                print o, ": [" + self.mingoal[o] + ", " + self.maxgoal[o]+"]"
        if self.user_add_mincom.get() != "N/A":
            print "Minimum Number of Commodities: " , self.user_add_mincom.get()
        if self.user_add_maxcom.get() != "N/A":
            print "Maximum Number of Commodities: " , self.user_add_maxcom.get()
        if self.user_cv_min.get() != "0":
            print "Minimum % C&V: " , self.user_cv_min.get()
        if self.user_cv_max.get() != "100":
            print "Maximum % C&V: " , self.user_cv_max.get()
        if self.supply_tact.get() != 0:
            print "Supply Tactical Demand: ", "Yes"
            for i in self.tactboxes.keys():
                if self.tactboxes[i].get() == 0:
                    print " > Excluded: ", i


    def draw_GUI(self,root):
        '''
        Draws the GUI and initialises all the windows
        '''

        # Layout settings
        self.bgcolor = '#%02x%02x%02x' % (51, 128, 255) # matches the logo background
        ttk.Style().configure(".", font=('Helvetica', 8), foreground="white") # applied to every widget
        ttk.Style().configure("TButton", background=self.bgcolor, foreground="black")
        ttk.Style().configure("TMenubutton", background=self.bgcolor) # applies to OptionMenu
        ttk.Style().configure("TLabel", background=self.bgcolor)
        ttk.Style().configure("TRadiobutton", background=self.bgcolor)
        ttk.Style().configure("TCheckbutton", background=self.bgcolor)
        ttk.Style().configure("TFrame", background=self.bgcolor)
        ttk.Style().configure("Treeview", background=self.bgcolor)
        ttk.Style().configure("Treeview.Heading", foreground="black")
        ttk.Style().configure("TEntry", background=self.bgcolor, foreground="black")

        # Create major GUI frames
        self.frame_left = ttk.Frame(root)
        self.frame_left.pack(side=LEFT, fill=Y, expand=1)
        self.breakline1 = ttk.Frame(root, relief=SUNKEN)
        self.breakline1.pack(side=LEFT, fill=Y, expand=1)
        self.frame_right = ttk.Frame(root)
        self.frame_right.pack(side=RIGHT, fill=BOTH, expand=1)
        self.padding0 = ttk.Frame(self.frame_left)
        self.padding0.pack(side=TOP, fill=Y, expand=1)
        self.frame_header = ttk.Frame(self.frame_left, padding=20)
        self.frame_header.pack(side=TOP)
        self.padding1a = ttk.Frame(self.frame_left)
        self.padding1a.pack(side=TOP, fill=Y, expand=1)
        self.breakline2 = ttk.Frame(self.frame_left, relief=SUNKEN)
        self.breakline2.pack(side=TOP, fill=X, expand=1)
        self.padding1b = ttk.Frame(self.frame_left)
        self.padding1b.pack(side=TOP, fill=Y, expand=1)
        self.frame_main = ttk.Frame(self.frame_left, padding=20)
        self.frame_main.pack(side=TOP)
        ttk.Label(self.frame_main, text = "            ").grid(row=0,column=4)
        self.padding2 = ttk.Frame(self.frame_left)
        self.padding2.pack(side=TOP, fill=Y, expand=1)
        self.frame_output = ttk.Frame(self.frame_right,padding=20)
        self.frame_output.pack(fill=BOTH, expand=1)

        # Fill the frames
        self.draw_header()
        self.draw_generalinput()
        self.draw_constraints()
        self.draw_fixfood()
        self.draw_editfood()
        self.draw_procurement()
        self.draw_routing()
        self.draw_CV()
        self.draw_funding()
        self.draw_obj()
        self.draw_save()
        self.draw_analysis()
        self.draw_outputs()
        self.draw_auto()
        self.draw_tact()
        self.draw_solution()
        self.draw_stdout()
        self.draw_scenlist()
        self.draw_activities()
        self.update_lists()

    def draw_header(self):
        '''
        Draws GUI component: Header
        '''

        ttk.Label(self.frame_header, text = "WFP's Assistant for Integrated Decision-Making", font = ("Helvetica", 16, "italic"), anchor=CENTER).grid(row=2,sticky=EW)
        ttk.Label(self.frame_header, text = "Food Basket | Sourcing | Delivery | Transfer Modality", font = ("Helvetica", 13), anchor=CENTER).grid(row=3,sticky=EW)
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'data')
        dest_dir = os.path.join(dest_dir, 'wfp2.gif')
        self.img = PhotoImage(file=dest_dir)
        # NB: http://ezgif.com/resize   is awesome for resizing gifs
        ttk.Label(self.frame_header, image = self.img, anchor=CENTER).grid(row=1,sticky=EW)

    def draw_generalinput(self):
        '''
        Draws GUI component: General user input
        '''

        r,c = 0,0
        ttk.Label(self.frame_main, text = "General User Input:", font = ("Helvetica",11,"bold"), anchor=CENTER).grid(row=r,column=c,columnspan=3,sticky=EW)
        label_hor = ttk.Label(self.frame_main, text = "Time Horizon")
        label_hor.grid(row=r+1,column=c,sticky=W)
        self.tstart = StringVar()
        self.tend = StringVar()
        ttk.OptionMenu(self.frame_main, self.tstart, self.periods[0], *self.periods).grid(row=r+1,column=c+1,sticky=EW)
        self.tstart.trace('w',self.update_lists)
        for t in self.periods:
            if sum(self.dem[b,t] for b in self.beneficiaries) > 0:
                p = t
                break
        ttk.OptionMenu(self.frame_main, self.tend, p, *self.periods).grid(row=r+1,column=c+2,sticky=EW)
        self.tend.trace('w',self.update_lists)
        t = ("The model will find the optimal decisions for each month in the time horizon."
            "\nIncreasing the amount of months considered increases the solution time."
            "\nAs a rule of thumb, adding 3 months doubles the solution time.")
        createToolTip(label_hor,t)
        self.horizon = list(self.periods) # contains all periods in the current time horizon
        for t in self.periods:
            if t != self.tstart.get():
                self.horizon.remove(t)
            else:
                break
        for t in reversed(self.periods):
            if t != self.tend.get():
                self.horizon.remove(t)
            else:
                break

        l = ttk.Label(self.frame_main, text = "Activity Type")
        l.grid(row=r+2,column=c,sticky=W)
        self.ben = StringVar()
        i = 0
        t = self.tend.get()
        m = self.dem[self.benlist[i], t]
        for b in self.benlist:
            if self.dem[b,t] > m :
                m = self.dem[b,t]
                i = self.benlist.index(b) # the default activity is the one with the biggest demand
        ttk.OptionMenu(self.frame_main, self.ben, self.benlist[i], *self.benlist).grid(row=r+2,column=c+1,columnspan=2,sticky=W)
        self.ben.trace('w',self.update_act)
        t = ("This is the activity or beneficiary type that the model will optimise."
            "\nNutritional performance is only tracked for this activity:"
            "\nOther activities may be supplied, but their nutrition is not tracked.")
        createToolTip(l,t)

        self.allowshortfalls = IntVar()
        self.allowshortfalls.set(0)
        label_sf = ttk.Label(self.frame_main, text = "Allow Shortfalls")
        label_sf.grid(row=r+3,column=c,sticky=W)
        ttk.Radiobutton(self.frame_main, text="Yes", variable = self.allowshortfalls, value=1).grid(row=r+3,column=c+1,sticky=W)
        ttk.Radiobutton(self.frame_main, text="No", variable=self.allowshortfalls, value=0).grid(row=r+3,column=c+2,sticky=W)
        t = ("By default the model has to supply 11 NVS to each beneficiary."
            "\nAllow shortfalls to investigate NVS-levels lower than 11.")
        createToolTip(label_sf,t)

        self.sensible = IntVar()
        self.sensible.set(1)
        label_sens = ttk.Label(self.frame_main, text = "Sensible Food Basket    ")
        label_sens.grid(row=r+4,column=c,sticky=W)
        ttk.Radiobutton(self.frame_main, text="Yes", variable=self.sensible, value=1).grid(row=r+4,column=c+1,sticky=W)
        ttk.Radiobutton(self.frame_main, text="No", variable=self.sensible, value=0).grid(row=r+4,column=c+2,sticky=W)
        t = ("This activates Sensible Food Basket constraints:"
            "\n1) 10-12% of Energy must be supplied through Proteins"
            "\n2) >17% of Energy must be supplied through Fats"
            "\n3) The food basket composition must follow these rules:"
            "\ng/p/d        <=            Food Group                       <=      g/p/d"
            "\n----------------------------------------------------------------------------------"
            "\n250 <=                (CEREALS & GRAINS)                         <= 500"
            "\n  30 <=                (PULSES & VEGETABLES)                    <= 130"
            "\n  15 <=                (OILS & FATS)                                    <=   40"
            "\n    5 <=                (IODISED SALT)  "
            "\n                            (MIXED & BLENDED FOODS)               <=   60"
            "\n                            (DAIRY PRODUCTS)                            <=   40"
            "\n                            (MEAT)                                                <=   40"
            "\n                            (FISH)                                                 <=   40")
        createToolTip(label_sens,t)
        # NB: The tooltip looks nicer in the GUI, for some reason the width of spaces in the GUI is about half the width of a letter

        self.useforecasts = IntVar()
        self.useforecasts.set(1)
        label_fc = ttk.Label(self.frame_main, text = "Use Forecasts")
        label_fc.grid(row=r+5,column=c,sticky=W)
        ttk.Radiobutton(self.frame_main, text="Yes", variable=self.useforecasts, value=1).grid(row=r+5,column=c+1,sticky=W)
        ttk.Radiobutton(self.frame_main, text="No", variable=self.useforecasts, value=0).grid(row=r+5,column=c+2,sticky=W)
        t = ("Prices will be adjusted based on seasonality forecasts."
            "\nThe price from the data is assumed to be the price in"
            "\nthe first month of the time horizon.")
        createToolTip(label_fc,t)

        self.supply_tact = IntVar()
        self.supply_tact.set(0)
##        label_td = ttk.Label(self.frame_main, text = "Tactical Demand")
##        label_td.grid(row=r+6,column=c,sticky=W)
        b_td = ttk.Button(self.frame_main, text = "Tactical Demand", command = lambda: self.show(self.tactwin))
        b_td.grid(row=r+6,column=c,sticky=W)
        ttk.Radiobutton(self.frame_main, text="Yes", variable=self.supply_tact, value=1).grid(row=r+6,column=c+1,sticky=W)
        ttk.Radiobutton(self.frame_main, text="No", variable=self.supply_tact, value=0).grid(row=r+6,column=c+2,sticky=W)
        t = ("Activate this option to supply Tactical Demand (as per inputs)."
            "\nTactical demand is supplied on top of any other activities.")
        createToolTip(b_td,t)

        l = ttk.Label(self.frame_main, text = "Activities To Supply")
        l.grid(row=r+7,column=c,sticky=W)
        self.act_button = ttk.Button(self.frame_main, text="Select (0)", command = lambda: self.show(self.actwin))
        self.act_button.grid(row=r+7,column=c+1,columnspan=2,sticky=EW)
        t = ("Select which activities you want to supply."
            "\n  (on top of the activity to be optimised)")
        createToolTip(l,t)

        self.varbasket = StringVar()
        label_vb = ttk.Label(self.frame_main, text = "Basket Variability")
        label_vb.grid(row=r+8,column=c,sticky=W)
        choices = ["Variable","Fix Commodities","Fix All"]
        ttk.OptionMenu(self.frame_main, self.varbasket, choices[0], *choices).grid(row=r+8,column=c+1,columnspan=2,sticky=W)
        t = ("Determines whether food baskets are allowed to change over time."
            "\nVariable = Commodities & Ration sizes may change over time."
            "\nFix Commodities = Ration sizes may change over time."
            "\nFix All = Commodities & Ration sizes remain constant.")
        createToolTip(label_vb,t)

        self.modality = StringVar()
        label_mod = ttk.Label(self.frame_main, text = "C&V Approach")
        label_mod.grid(row=r+9,column=c,sticky=W)
        mods = ["Cash", "Voucher"]
        ttk.OptionMenu(self.frame_main, self.modality, mods[1], *mods).grid(row=r+9,column=c+1,columnspan=2,sticky=W)
        t = ("Determines how the model approaches the C&V Transfer Modality."
            "\nVoucher = We provide vouchers for specific commodities."
            "\nCash = We provide a cash contribution, which is then spent"
            "\naccording to a pre-defined expenditure pattern.")
        createToolTip(label_mod,t)

        ttk.Label(self.frame_main, text = " ").grid(row=r+10,column=c,sticky=W)

    def draw_constraints(self):
        '''
        Draws GUI component: Specific user input
        '''

        r,c = 11,0
        ttk.Label(self.frame_main, text = "Specific User Input:", font = ("Helvetica",11,"bold"), anchor=CENTER).grid(row=r,column=c,columnspan=3,sticky=EW)
        ttk.Button(self.frame_main, text = "Edit Food Basket", command = lambda: self.show(self.foodwin)).grid(row=r+1,column=c,sticky=EW)
        ttk.Button(self.frame_main, text = "Fix Food Basket", command = lambda: self.show(self.fixwin)).grid(row=r+1,column=c+1,columnspan=2,sticky=EW)
        ttk.Button(self.frame_main, text = "Procurement", command = lambda: self.show(self.procwin)).grid(row=r+2,column=c,sticky=EW)
        ttk.Button(self.frame_main, text = "Cash & Vouchers", command = lambda: self.show(self.cvwin)).grid(row=r+2,column=c+1,columnspan=2,sticky=EW)
        ttk.Button(self.frame_main, text = "Logistics", command = lambda: self.show(self.routewin)).grid(row=r+3,column=c,sticky=EW)
        ttk.Button(self.frame_main, text = "Objectives", command = lambda: self.show(self.objwin)).grid(row=r+3,column=c+1,columnspan=2,sticky=EW)
        ttk.Button(self.frame_main, text = "Funding & In-Kind Donations", command = lambda: self.show(self.funwin)).grid(row=r+4,column=c,columnspan=3,sticky=EW)

    def draw_fixfood(self):
        '''
        Draws GUI component: Pop-up window for fixed food baskets
        '''

        self.food2fix = []
        self.fixwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.fixwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.fixwin)) # This overrides the red 'x' with a new exit command: return to main window
        ttk.Label(self.fixwin, text = "Fix Food Basket",font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky = W,pady=(0,5))
        ttk.Label(self.fixwin, text = "Commodity Type", anchor=CENTER).grid(row=1,column=0)
        ttk.Label(self.fixwin, text = "Specific Commodity", anchor=CENTER).grid(row=1,column=1)
        ttk.Label(self.fixwin, text = "Quantity (g/p/d)", anchor=CENTER).grid(row=1,column=2)
        self.fix_spec={}
        self.fix_com={}
        self.fix_list={}
        self.fix_quant={}
        temp = {}
        for i in range (0,15):
            temp[i] = int(i)
            self.fix_spec[i] = StringVar()
            ttk.OptionMenu(self.fixwin, self.fix_spec[i], "Filter", *self.supcom).grid(row=2+i,column=0,sticky=EW,padx=(5,0))
            self.fix_spec[i].trace('w', self.update_fix_com)
            self.fix_com[i] = StringVar()
            self.fix_list[i]=ttk.OptionMenu(self.fixwin, self.fix_com[i], "Select", *self.commodities)
            self.fix_list[i].grid(row=2+i,column=1,sticky=EW)
            self.fix_quant[i] = StringVar()
            self.fix_quant[i].set("N/A")
            ttk.Entry(self.fixwin, textvariable = self.fix_quant[i], justify=CENTER).grid(row=2+i,column=2,sticky=EW)
        self.fix_spec[0].trace('w', self.getbaseline)
        self.fix_spec[0].set("test")
        self.fix_spec[0].set("Filter")
        # NB: I can't add the 'i' value to the update_fix_com call, but I can obtain the appropriate 'i' value with this workaround

        ttk.Label(self.fixwin, text = "          ").grid(row=0,column=4,sticky=W)
        ttk.Label(self.fixwin, text = "Conversion Table",font=("Helvetica",11,"bold")).grid(row=0,column=5,columnspan=3,sticky=W,pady=(0,5))
        ttk.Label(self.fixwin, text = "Feeding Days").grid(row=2, column=5, sticky=EW)
        ttk.Label(self.fixwin, text = "Household Size").grid(row=3, column=5, sticky=EW)
        ttk.Label(self.fixwin, text = "Conversion g/L").grid(row=4, column=5, sticky=EW)
        self.fix_days = DoubleVar()
        self.fix_days.set(max(self.feedingdays[self.ben.get(),k] for k in self.commodities))
        self.fix_days.trace('w',self.update_conversion)
        ttk.Entry(self.fixwin, textvariable=self.fix_days, justify=CENTER).grid(row=2, column=6, sticky=EW)
        self.fix_hh = DoubleVar()
        self.fix_hh.set(5)
        self.fix_hh.trace('w',self.update_conversion)
        ttk.Entry(self.fixwin, textvariable=self.fix_hh, justify=CENTER).grid(row=3, column=6, sticky=EW)
        self.fix_g2l = DoubleVar()
        self.fix_g2l.set(910)
        self.fix_g2l.trace('w',self.update_conversion)
        ttk.Entry(self.fixwin, textvariable=self.fix_g2l, justify=CENTER).grid(row=4, column=6, sticky=EW)
        ttk.Label(self.fixwin, text = "g/p/d", anchor=CENTER).grid(row=1, column=9)
        self.fix_def = {}
        self.fix_def[1,1] = ttk.Label(self.fixwin, text="10 KG", anchor=E)
        self.fix_def[1,1].grid(row=2, column=8, sticky=EW)
        self.fix_def[1,2] = ttk.Label(self.fixwin, text="", anchor=E)
        self.fix_def[1,2].grid(row=2, column=9, sticky=EW)
        self.fix_def[2,1] = ttk.Label(self.fixwin, text="25 KG", anchor=E)
        self.fix_def[2,1].grid(row=3, column=8, sticky=EW)
        self.fix_def[2,2] = ttk.Label(self.fixwin, text="", anchor=E)
        self.fix_def[2,2].grid(row=3, column=9, sticky=EW)
        self.fix_def[3,1] = ttk.Label(self.fixwin, text="50 KG", anchor=E)
        self.fix_def[3,1].grid(row=4, column=8, sticky=EW)
        self.fix_def[3,2] = ttk.Label(self.fixwin, text="", anchor=E)
        self.fix_def[3,2].grid(row=4, column=9, sticky=EW)
        self.fix_def[4,1] = ttk.Label(self.fixwin, text="1    L", anchor=E)
        self.fix_def[4,1].grid(row=5, column=8, sticky=EW)
        self.fix_def[4,2] = ttk.Label(self.fixwin, text="", anchor=E)
        self.fix_def[4,2].grid(row=5, column=9, sticky=EW)
        self.fix_def[5,1] = ttk.Label(self.fixwin, text="5    L", anchor=E)
        self.fix_def[5,1].grid(row=6, column=8, sticky=EW)
        self.fix_def[5,2] = ttk.Label(self.fixwin, text="", anchor=E)
        self.fix_def[5,2].grid(row=6, column=9, sticky=EW)
        ttk.Label(self.fixwin, text = " ").grid(row=7,column=5,sticky=W)
        self.fix_in = DoubleVar()
        self.fix_in.set(3.14)
        self.fix_in.trace('w',self.update_conversion)
        ttk.Entry(self.fixwin, textvariable=self.fix_in, justify=CENTER).grid(row=8, column=5, sticky=EW)
        self.fix_unit1 = StringVar()
        ttk.OptionMenu(self.fixwin, self.fix_unit1, "g/p/d", *["g/p/d","KG/hh/m","L/hh/m"]).grid(row=8, column=6, sticky=EW)
        self.fix_unit1.trace('w',self.update_conversion)
        ttk.Label(self.fixwin, text = "        =        ", anchor=CENTER).grid(row=8, column=7, sticky=EW)
        self.fix_out = ttk.Label(self.fixwin, text = "", anchor=CENTER)
        self.fix_out.grid(row=8, column=8, sticky=EW)
        self.fix_unit2 = StringVar()
        ttk.OptionMenu(self.fixwin, self.fix_unit2, "KG/hh/m", *["g/p/d","KG/hh/m","L/hh/m"]).grid(row=8, column=9, sticky=EW)
        self.fix_unit2.trace('w',self.update_conversion)
        self.update_conversion()

        ttk.Label(self.fixwin, text = " ").grid(row=17,column=0,sticky=W)
        ttk.Button(self.fixwin, text="Load from data", command = self.load_basket).grid(row=18,column=1,sticky=EW)
        ttk.Button(self.fixwin, text="Add constraints", command = self.set_fb).grid(row=18,column=2,sticky=EW)
        ttk.Button(self.fixwin, text="Reset", command = self.reset_fix).grid(row=19,column=1,sticky=EW)
        ttk.Button(self.fixwin, text="Back", command = lambda: self.close(self.fixwin)).grid(row=19,column=2,sticky=EW)
        self.fixwin.withdraw()

    def draw_editfood(self):
        '''
        Draws GUI component: Pop-up window for editing food basket decisions
        '''

        self.user_ex_com = []
        self.user_add_nut = {}
        self.user_add_fg = {}
        self.user_ex_fg = []
        self.foodwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.foodwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.foodwin))
        ttk.Label(self.foodwin, text = "Commodity Decisions",font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky=W,pady=(0,5))
        ttk.Label(self.foodwin, text = "Include", anchor=W).grid(row=2,column=0,sticky=W)
        ttk.Label(self.foodwin, text = "Commodity Type", anchor=CENTER).grid(row=1,column=1)
        ttk.Label(self.foodwin, text = "Specific Commodity", anchor=CENTER).grid(row=1,column=2)
        ttk.Label(self.foodwin, text = "Minimum (g/p/d)", anchor=CENTER).grid(row=1,column=3)
        ttk.Label(self.foodwin, text = "Maximum (g/p/d)", anchor=CENTER).grid(row=1,column=4)
        self.fb_add_com = StringVar()
        ttk.OptionMenu(self.foodwin, self.fb_add_com, "Filter", *self.supcom).grid(row = 2,column = 1,sticky=EW)
        self.fb_add_com.trace('w',self.update_fb_add_speccom) # Runs self.update_fb_add_speccom when comchoice changes
        self.fb_add_speccom = StringVar()
        self.fb_add_list = ttk.OptionMenu(self.foodwin, self.fb_add_speccom, "Select", *self.commodities) # OptionMenu is filled through self.update_fb_add_speccom
        self.fb_add_list.grid(row=2,column=2,sticky=EW)
        self.fb_add_min = StringVar()
        self.fb_add_min.set("N/A")
        self.fb_add_max = StringVar()
        self.fb_add_max.set("N/A")
        ttk.Entry(self.foodwin, textvariable = self.fb_add_min, justify=CENTER).grid(row=2,column=3,sticky = EW)
        ttk.Entry(self.foodwin, textvariable = self.fb_add_max, justify=CENTER).grid(row=2,column=4,sticky = EW)
        ttk.Button(self.foodwin,text="Add constraint", command = self.include_com).grid(row=2,column=5,sticky = EW)

        ttk.Label(self.foodwin, text = "Exclude", anchor=W).grid(row=3,column=0,sticky=W)
        self.fb_ex_com = StringVar()
        self.fb_ex_speccom = StringVar()
        ttk.OptionMenu(self.foodwin, self.fb_ex_com, "Filter", *self.supcom).grid(row = 3,column = 1,sticky = EW)
        self.fb_ex_com.trace('w',self.update_fb_ex_speccom)
        self.fb_ex_list = ttk.OptionMenu(self.foodwin, self.fb_ex_speccom, "Select", *self.commodities)
        self.fb_ex_list.grid(row=3,column=2,sticky=EW)
        ttk.Button(self.foodwin,text = "Add constraint", command = self.exclude_com).grid(row = 3, column = 5,sticky = EW)

        ttk.Label(self.foodwin, text = " ").grid(row = 4, column = 0,sticky = W)
        ttk.Label(self.foodwin, text = "Food Group Decisions",font=("Helvetica",11,"bold")).grid(row=5,column=0,columnspan=3,sticky = W,pady=(0,5))
        ttk.Label(self.foodwin, text = "Include", anchor=W).grid(row=7,column=0,sticky=W)
        ttk.Label(self.foodwin, text = "Food group", anchor=CENTER).grid(row=6,column=1)
        ttk.Label(self.foodwin, text = "Minimum (g/p/d)", anchor=CENTER).grid(row=6,column=2)
        ttk.Label(self.foodwin, text = "Maximum (g/p/d)", anchor=CENTER).grid(row=6,column=3)
        self.fb_add_fg = StringVar()
        ttk.OptionMenu(self.foodwin, self.fb_add_fg, "Select", *self.foodgroups).grid(row=7,column=1,sticky=EW)
        self.fb_add_fg_min = StringVar()
        self.fb_add_fg_min.set("N/A")
        ttk.Entry(self.foodwin, textvariable=self.fb_add_fg_min, justify=CENTER).grid(row=7,column=2,sticky=EW)
        self.fb_add_fg_max = StringVar()
        self.fb_add_fg_max.set("N/A")
        ttk.Entry(self.foodwin, textvariable=self.fb_add_fg_max, justify=CENTER).grid(row=7,column=3,sticky=EW)
        ttk.Button(self.foodwin, text = "Add constraint", command=self.include_fg).grid(row=7,column=5,sticky=EW)

        ttk.Label(self.foodwin, text = "Exclude", anchor=W).grid(row=8,column=0,sticky=W)
        self.fb_ex_fg = StringVar()
        ttk.OptionMenu(self.foodwin, self.fb_ex_fg, "Select", *self.foodgroups).grid(row=8,column=1,sticky=EW)
        ttk.Button(self.foodwin, text = "Add constraint", command=self.exclude_fg).grid(row=8,column=5,sticky=EW)

        ttk.Label(self.foodwin, text = " ").grid(row = 9, column = 0,sticky = W)
        ttk.Label(self.foodwin, text = "Commodity Diversification",font=("Helvetica",11,"bold")).grid(row=10,column=0,columnspan=3,sticky = W,pady=(0,5))
        ttk.Label(self.foodwin, text = "Minimum #", anchor=W).grid(row = 11, column = 0,sticky=W)
        ttk.Label(self.foodwin, text = "Maximum #", anchor=W).grid(row = 12, column = 0,sticky=W)
        self.user_add_mincom = StringVar()
        self.user_add_mincom.set("N/A")
        ttk.Entry(self.foodwin,textvariable=self.user_add_mincom, justify=CENTER).grid(row = 11, column = 1)
        self.user_add_maxcom = StringVar()
        self.user_add_maxcom.set("N/A")
        ttk.Entry(self.foodwin,textvariable=self.user_add_maxcom, justify=CENTER).grid(row = 12, column = 1)

        ttk.Label(self.foodwin, text = " ").grid(row = 13, column = 0, sticky = W)
        ttk.Label(self.foodwin, text = "Nutritient Restrictions",font=("Helvetica",11,"bold")).grid(row=14,column=0,columnspan=3,sticky = W,pady=(0,5))
        ttk.Label(self.foodwin, text = "Limit Shortfall", anchor=W).grid(row=16,column=0,sticky=W)
        ttk.Label(self.foodwin, text = "Nutrient", anchor=CENTER).grid(row=15,column=1)
        ttk.Label(self.foodwin, text = "Maximum shortfall (%)", anchor=CENTER).grid(row=15,column=2)
        self.nutchoice = StringVar()
        nutlist = list(self.nutrients)
        nutlist.append("All")
        ttk.OptionMenu(self.foodwin, self.nutchoice, "Select", *nutlist).grid(row=16,column=1,sticky=EW)
        self.maxnut = StringVar()
        self.maxnut.set("N/A")
        ttk.Entry(self.foodwin, textvariable=self.maxnut, justify=CENTER).grid(row=16,column=2,sticky=EW)
        ttk.Button(self.foodwin,text = "Add constraint",command = self.set_nut).grid(row=16,column=5,sticky=EW)
        ttk.Label(self.foodwin, text = "Minimum (%)", anchor=CENTER).grid(row=17,column=1)
        ttk.Label(self.foodwin, text = "Maximum (%)", anchor=CENTER).grid(row=17,column=2)
        ttk.Label(self.foodwin, text = "Energy from Proteins", anchor=W).grid(row=18,column=0,sticky=W)
        self.user_nut_minprot = StringVar()
        self.user_nut_minprot.set("0")
        ttk.Entry(self.foodwin,textvariable=self.user_nut_minprot, justify=CENTER).grid(row = 18, column = 1)
        self.user_nut_maxprot = StringVar()
        self.user_nut_maxprot.set("100")
        ttk.Entry(self.foodwin,textvariable=self.user_nut_maxprot, justify=CENTER).grid(row = 18, column = 2)
        ttk.Label(self.foodwin, text = "Energy from Fats", anchor=W).grid(row=19,column=0,sticky=W)
        self.user_nut_minfat = StringVar()
        self.user_nut_minfat.set("0")
        ttk.Entry(self.foodwin,textvariable=self.user_nut_minfat, justify=CENTER).grid(row = 19, column = 1)
        self.user_nut_maxfat = StringVar()
        self.user_nut_maxfat.set("100")
        ttk.Entry(self.foodwin,textvariable=self.user_nut_maxfat, justify=CENTER).grid(row = 19, column = 2)

        ttk.Label(self.foodwin,text=" ").grid(row=20,column=0,sticky=EW)
        ttk.Label(self.foodwin, text = "Acceptance of GMOs",font=("Helvetica",11,"bold")).grid(row=21,column=0,columnspan=3,sticky = W,pady=(0,5))
        self.gmo = IntVar()
        self.gmo.set(1)
        ttk.Radiobutton(self.foodwin, text = "Allowed", variable=self.gmo, value=1).grid(row=22, column=0, sticky=EW)
        ttk.Radiobutton(self.foodwin, text = "Not allowed", variable=self.gmo, value=0).grid(row=22, column=1,sticky=EW)

        ttk.Label(self.foodwin,text=" ").grid(row=23,column=0,sticky=EW)
        ttk.Label(self.foodwin,text=" ").grid(row=0,column=6,sticky=EW)
        ttk.Button(self.foodwin, text = "Reset", command = self.reset_fb,width=10).grid(row=24, column=4)
        ttk.Button(self.foodwin, text = "Back", command = lambda: self.close(self.foodwin),width=10).grid(row=24, column=5)
        self.foodwin.withdraw()

    def draw_procurement(self):
        '''
        Draws GUI component: Pop-up window for editing procurement decisions
        '''

        self.user_ex_proc_int = {}
        self.user_ex_proc_loc = {}
        self.user_add_proc_int = {}
        self.user_add_proc_loc = {}
        self.procwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.procwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.procwin))

        ttk.Label(self.procwin, text = "International Procurement Decisions:",font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky=W,pady=(0,5))
        ttk.Label(self.procwin, text = "Include").grid(row=2,column=0,sticky = W)
        ttk.Label(self.procwin, text = "Origin Country", anchor=CENTER).grid(row=1,column=1)
        self.proc_add_src_int= StringVar()
        countrylist = [] # reduced list
        incolist = []
        for key in self.proccap.keys(): # key = (src,ndp,com)
            if key[1] not in (self.LSs+self.LMs):
                countrylist.append(key[0][:-6])
                incolist.append(key[0][-3:])
        countrylist.append("Any")
        incolist.append("Any")
        countrylist = list(set(countrylist))
        countrylist.sort()
        incolist = list(set(incolist))
        incolist.sort()
        ttk.OptionMenu(self.procwin, self.proc_add_src_int, "Select", *countrylist).grid(row = 2,column = 1,sticky=EW)
        self.proc_add_src_int.trace('w',self.update_proc_add_inco_int)
        ttk.Label(self.procwin,text="Incoterm", anchor=CENTER).grid(row=1,column=2)
        self.proc_add_inco_int= StringVar()
        ttk.OptionMenu(self.procwin, self.proc_add_inco_int, "Select", *incolist).grid(row=2,column=2,sticky=EW)
        self.proc_add_inco_int.trace('w',self.update_proc_add_ndp_int) # Runs self.update_proc_add_ndp_int when value changes (i.e. when an incoterm is selected)
        ttk.Label(self.procwin, text = "Named Delivery Place", anchor=CENTER).grid(row=1,column=3)
        self.proc_add_ndp_int = StringVar()
        self.proc_add_list_ndp_int = ttk.OptionMenu(self.procwin, self.proc_add_ndp_int, "Select", *self.ISs+self.RSs)
        self.proc_add_list_ndp_int.grid(row=2,column=3,sticky=EW)
        self.proc_add_ndp_int.trace('w',self.update_proc_add_com_int)
        ttk.Label(self.procwin, text = "Commodity", anchor=CENTER).grid(row=1,column=4)
        self.proc_add_com_int= StringVar()
        self.proc_add_list_com_int = ttk.OptionMenu(self.procwin, self.proc_add_com_int, "Select", *self.commodities)
        self.proc_add_list_com_int.grid(row = 2,column = 4,sticky=EW)
        self.proc_add_mt_int = StringVar()
        self.proc_add_mt_int.set("N/A")
        ttk.Label(self.procwin, text = "Quantity per Month (mt)", anchor=CENTER).grid(row=1,column=5)
        ttk.Entry(self.procwin, textvariable = self.proc_add_mt_int, justify=CENTER).grid(row=2,column=5,sticky = EW)
        ttk.Button(self.procwin,text="Add constraint", command = self.include_proc_int).grid(row=2,column=6,sticky = EW)

        ttk.Label(self.procwin, text = "Exclude").grid(row=3,column=0,sticky = W)
        self.proc_ex_src_int= StringVar()
        ttk.OptionMenu(self.procwin, self.proc_ex_src_int, "Select", *countrylist).grid(row = 3,column = 1, sticky = EW)
        self.proc_ex_src_int.trace('w',self.update_proc_ex_ndp_int)
        self.proc_ex_ndp_int = StringVar()
        self.proc_ex_list_ndp_int = ttk.OptionMenu(self.procwin, self.proc_ex_ndp_int, "Select", *self.ISs+self.RSs)
        self.proc_ex_ndp_int.trace('w',self.update_proc_ex_com_int)
        self.proc_ex_list_ndp_int.grid(row=3,column=3,sticky=EW)
        self.proc_ex_com_int= StringVar()
        self.proc_ex_list_com_int = ttk.OptionMenu(self.procwin, self.proc_ex_com_int, "Select", *self.commodities)
        self.proc_ex_list_com_int.grid(row=3,column=4,sticky=EW)
        ttk.Button(self.procwin,text = "Add constraint", command = self.exclude_proc_int).grid(row = 3, column = 6)

        ttk.Label(self.procwin, text = " ").grid(row=4,column=0,sticky = W)
        ttk.Label(self.procwin, text = "Local Procurement Decisions",font=("Helvetica",11,"bold")).grid(row=5,column=0,columnspan=3,sticky = W,pady=(0,5))
        ttk.Label(self.procwin, text = "Include").grid(row=7,column=0,sticky = W)
        ttk.Label(self.procwin, text = "Origin Country", anchor=CENTER).grid(row=6,column=1)
        self.proc_add_src_loc= StringVar()
        countrylist = [] # reduced list
        incolist = []
        for key in self.proccap.keys(): # key = (src,ndp,com)
            if key[1] in self.LSs:
                countrylist.append(key[0][:-6])
                incolist.append(key[0][-3:])
        countrylist.append("Any")
        incolist.append("Any")
        countrylist = list(set(countrylist))
        countrylist.sort()
        incolist = list(set(incolist))
        incolist.sort()
        ttk.OptionMenu(self.procwin, self.proc_add_src_loc, "Select", *countrylist).grid(row = 7,column = 1,sticky=EW)
        self.proc_add_src_loc.trace('w',self.update_proc_add_inco_loc)
        ttk.Label(self.procwin,text="Incoterm", anchor=CENTER).grid(row=6,column=2)
        self.proc_add_inco_loc= StringVar()
        ttk.OptionMenu(self.procwin, self.proc_add_inco_loc, "Select", *incolist).grid(row=7,column=2,sticky=EW)
        self.proc_add_inco_loc.trace('w',self.update_proc_add_ndp_loc)
        ttk.Label(self.procwin, text = "Named Delivery Place:", anchor=CENTER).grid(row=6,column=3)
        self.proc_add_ndp_loc = StringVar()
        self.proc_add_list_ndp_loc = ttk.OptionMenu(self.procwin, self.proc_add_ndp_loc, "Select", *self.LSs)
        self.proc_add_ndp_loc.trace('w',self.update_proc_add_com_loc)
        self.proc_add_list_ndp_loc.grid(row=7,column=3,sticky=EW)
        ttk.Label(self.procwin, text = "Commodity:", anchor=CENTER).grid(row=6,column=4)
        self.proc_add_com_loc= StringVar()
        self.proc_add_list_com_loc = ttk.OptionMenu(self.procwin, self.proc_add_com_loc, "Select", *self.commodities)
        self.proc_add_list_com_loc.grid(row = 7,column = 4,sticky=EW)
        self.proc_add_mt_loc = StringVar()
        self.proc_add_mt_loc.set("N/A")
        ttk.Label(self.procwin, text = "Quantity (mt)", anchor=CENTER).grid(row=6,column=5)
        ttk.Entry(self.procwin, textvariable = self.proc_add_mt_loc, justify=CENTER).grid(row=7,column=5,sticky = EW)
        ttk.Button(self.procwin,text="Add constraint", command = self.include_proc_loc).grid(row=7,column=6,sticky = EW)

        ttk.Label(self.procwin, text = "Exclude").grid(row=8,column=0,sticky = W)
        self.proc_ex_src_loc= StringVar()
        ttk.OptionMenu(self.procwin, self.proc_ex_src_loc, "Select", *countrylist).grid(row = 8,column = 1, sticky = EW)
        self.proc_ex_src_loc.trace('w',self.update_proc_ex_ndp_loc)
        self.proc_ex_ndp_loc = StringVar()
        self.proc_ex_list_ndp_loc = ttk.OptionMenu(self.procwin, self.proc_ex_ndp_loc, "Select", *self.LSs)
        self.proc_ex_ndp_loc.trace('w',self.update_proc_ex_com_loc)
        self.proc_ex_list_ndp_loc.grid(row=8,column=3,sticky=EW)
        self.proc_ex_com_loc= StringVar()
        self.proc_ex_list_com_loc = ttk.OptionMenu(self.procwin, self.proc_ex_com_loc, "Select", *self.commodities)
        self.proc_ex_list_com_loc.grid(row=8,column=4,sticky=EW)
        ttk.Button(self.procwin,text = "Add constraint", command = self.exclude_proc_loc).grid(row = 8, column = 6)

        ttk.Label(self.procwin, text = " " ).grid(row=13,column=0,sticky=W)
        ttk.Label(self.procwin, text = "Procurement Allocation",font=("Helvetica",11,"bold")).grid(row=14, column=0,columnspan=3, sticky=W,pady=(0,5))
        ttk.Label(self.procwin, text = "Min %", anchor=CENTER).grid(row=15,column=2)
        ttk.Label(self.procwin, text = "Max %", anchor=CENTER).grid(row=15,column=3)
        ttk.Label(self.procwin, text = "International Procurement").grid(row=16,column=0,columnspan=2,sticky=W)
        self.user_int_min = StringVar()
        self.user_int_min.set("0")
        self.user_int_max = StringVar()
        self.user_int_max.set("100")
        ttk.Entry(self.procwin, textvariable = self.user_int_min, justify=CENTER).grid(row=16, column=2, sticky=EW)
        ttk.Entry(self.procwin, textvariable = self.user_int_max, justify=CENTER).grid(row=16, column=3, sticky=EW)
        ttk.Label(self.procwin, text = "Regional Procurement").grid(row=17,column=0,columnspan=2,sticky=W)
        self.user_reg_min = StringVar()
        self.user_reg_min.set("0")
        self.user_reg_max = StringVar()
        self.user_reg_max.set("100")
        ttk.Entry(self.procwin, textvariable = self.user_reg_min, justify=CENTER).grid(row=17, column=2, sticky=EW)
        ttk.Entry(self.procwin, textvariable = self.user_reg_max, justify=CENTER).grid(row=17, column=3, sticky=EW)
        ttk.Label(self.procwin, text = "Local Procurement").grid(row=18,column=0,columnspan=2,sticky=W)
        self.user_loc_min = StringVar()
        self.user_loc_min.set("0")
        self.user_loc_max = StringVar()
        self.user_loc_max.set("100")
        ttk.Entry(self.procwin, textvariable = self.user_loc_min, justify=CENTER).grid(row=18, column=2, sticky=EW)
        ttk.Entry(self.procwin, textvariable = self.user_loc_max, justify=CENTER).grid(row=18, column=3, sticky=EW)

        ttk.Label(self.procwin, text = " ").grid(row=19,column=0)
        ttk.Button(self.procwin, text = "Reset", command = self.reset_proc,width=10).grid(row=20, column=6, sticky=E)
        ttk.Button(self.procwin, text = "Back", command = lambda: self.close(self.procwin),width=10).grid(row=20, column=8, sticky=W)
        self.procwin.withdraw()

        ttk.Label(self.procwin, text = "   ").grid(row=0,column=7,sticky=EW)
        ttk.Label(self.procwin, text = "   ").grid(row=0,column=9,sticky=EW)
        ttk.Label(self.procwin, text = "Select months:").grid(row=0,column=8,sticky=W)
        self.listboxP = Listbox(self.procwin, height=16,width=10, selectmode=EXTENDED, exportselection=FALSE)
        self.listboxP.grid(row=1,rowspan=10,column=8,sticky=EW)
        for t in self.periods:
            self.listboxP.insert(END,t)

    def draw_routing(self):
        '''
        Draws GUI component: Pop-up window for editing routing decisions
        '''

        self.user_ex_route = {}
        self.user_add_route = {}
        self.user_cap_util = {}
        self.user_cap_aloc = {}
        self.routewin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.routewin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.routewin))
        ttk.Label(self.routewin, text = "Routing Decisions",font = ("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky = W,pady=(0,5))
        ttk.Label(self.routewin, text = "Include").grid(row=2,column=0,sticky=W)
        ttk.Label(self.routewin, text = "Origin Type", anchor=CENTER).grid(row=1,column=1)
        self.route_add_type = StringVar()
        loctypes = ["Load Port","Discharge Port","Extended Distribution Point","Regional Market","Local Market","Local Supplier"]
        ttk.OptionMenu(self.routewin, self.route_add_type, "Select", *loctypes).grid(row=2,column=1,sticky=EW)
        self.route_add_type.trace('w',self.update_route_add_loc1)
        ttk.Label(self.routewin, text = "Origin", anchor=CENTER).grid(row=1,column=2)
        self.route_add_loc1 = StringVar()
        self.route_add_list_loc1 = ttk.OptionMenu(self.routewin, self.route_add_loc1, "Select", *(self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs))
        self.route_add_list_loc1.grid(row=2,column=2,sticky=EW)
        self.route_add_loc1.trace('w',self.update_route_add_loc2)
        ttk.Label(self.routewin, text = "Destination", anchor=CENTER).grid(row=1,column=3)
        self.route_add_loc2 = StringVar()
        self.route_add_list_loc2 = ttk.OptionMenu(self.routewin, self.route_add_loc2, "Select", *(self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs))
        self.route_add_list_loc2.grid(row=2,column=3,sticky=EW)
        ttk.Label(self.routewin, text = "Commodity", anchor=CENTER).grid(row=1,column=4)
        self.route_add_com = StringVar()
        ttk.OptionMenu(self.routewin, self.route_add_com, "Select", *self.commodities).grid(row=2,column=4,sticky=EW)
        ttk.Label(self.routewin, text = "Quantity (mt)", anchor=CENTER).grid(row=1,column=5)
        self.route_add_mt = StringVar()
        self.route_add_mt.set("N/A")
        ttk.Entry(self.routewin, textvariable = self.route_add_mt, justify=CENTER).grid(row=2,column=5,sticky = EW)
        ttk.Button(self.routewin,text="Add constraint", command = self.include_route).grid(row=2,column=6,sticky = EW)

        ttk.Label(self.routewin, text = "Exclude").grid(row=3,column=0,sticky = W)
        self.route_ex_type = StringVar()
        ttk.OptionMenu(self.routewin, self.route_ex_type, "Select", *loctypes).grid(row=3,column=1,sticky=EW)
        self.route_ex_type.trace('w',self.update_route_ex_loc1)
        self.route_ex_loc1 = StringVar()
        self.route_ex_list_loc1 = ttk.OptionMenu(self.routewin, self.route_ex_loc1, "Select", *(self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs))
        self.route_ex_list_loc1.grid(row=3,column=2,sticky=EW)
        self.route_ex_loc1.trace('w',self.update_route_ex_loc2)
        self.route_ex_loc2 = StringVar()
        self.route_ex_list_loc2 = ttk.OptionMenu(self.routewin, self.route_ex_loc2, "Select", *(self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs))
        self.route_ex_list_loc2.grid(row=3,column=3,sticky=EW)
        self.route_ex_com = StringVar()
        comlist = list(self.commodities)
        comlist.append("Any")
        ttk.OptionMenu(self.routewin, self.route_ex_com, "Any", *comlist).grid(row=3,column=4,sticky=EW)
        ttk.Button(self.routewin,text = "Add constraint", command = self.exclude_route).grid(row = 3, column = 6)

        ttk.Label(self.routewin, text = " ").grid(row=4,column=0)
        ttk.Label(self.routewin, text = "Capacity Decisions",font = ("Helvetica",11,"bold")).grid(row=5,column=0,columnspan=3,sticky = W,pady=(0,5))
        l = ttk.Label(self.routewin, text = "Utilisation")
        l.grid(row=7,column=0,sticky=W)
        t = ("A utilisation of 80% means that 80% of the available capacity is used"
            "\nThis is used when the available capacity needs to be shared with other COs")
        createToolTip(l,t)
        l = ttk.Label(self.routewin, text = "Allocation")
        l.grid(row=8,column=0,sticky=W)
        t = ("An allocation of 30% for 'DP 1' means that 30% of all metric tonnes moving through DPs move through 'DP 1'"
            "\nThis is used to spread out the incoming commodities over different DPs and EDPs (preventing congestion)")
        createToolTip(l,t)
        ttk.Label(self.routewin, text = "Location", anchor=CENTER).grid(row=6,column=1)
        self.route_util_loc = StringVar()
        self.route_aloc_loc = StringVar()
        ttk.OptionMenu(self.routewin, self.route_util_loc, "Select", *(self.DPs+self.EDPs)).grid(row=7,column=1,sticky=EW)
        ttk.OptionMenu(self.routewin, self.route_aloc_loc, "Select", *(self.DPs+self.EDPs)).grid(row=8,column=1,sticky=EW)
        ttk.Label(self.routewin, text = "Min %", anchor=CENTER).grid(row=6,column=2)
        ttk.Label(self.routewin, text = "Max %", anchor=CENTER).grid(row=6,column=3)
        self.route_util_min = StringVar()
        self.route_util_max = StringVar()
        self.route_aloc_min = StringVar()
        self.route_aloc_max = StringVar()
        self.route_util_min.set("0")
        self.route_util_max.set("100")
        self.route_aloc_min.set("0")
        self.route_aloc_max.set("100")
        ttk.Entry(self.routewin, text = self.route_util_min, justify=CENTER).grid(row=7,column=2,sticky=EW)
        ttk.Entry(self.routewin, text = self.route_util_max, justify=CENTER).grid(row=7,column=3,sticky=EW)
        ttk.Entry(self.routewin, text = self.route_aloc_min, justify=CENTER).grid(row=8,column=2,sticky=EW)
        ttk.Entry(self.routewin, text = self.route_aloc_max, justify=CENTER).grid(row=8,column=3,sticky=EW)
        ttk.Button(self.routewin, text = "Add constraint", command = self.set_cap_util).grid(row=7,column=4,sticky=EW)
        ttk.Button(self.routewin, text = "Add constraint", command = self.set_cap_aloc).grid(row=8,column=4,sticky=EW)

        ttk.Label(self.routewin, text = " ").grid(row = 9, column = 0)
        ttk.Button(self.routewin, text = "Reset", command = self.reset_route,width=10).grid(row=10, column=6, sticky=E)
        ttk.Button(self.routewin, text = "Back", command = lambda: self.close(self.routewin),width=10).grid(row=10, column=8)

        ttk.Label(self.routewin, text = "   ").grid(row=0,column=7,sticky=EW)
        ttk.Label(self.routewin, text = "   ").grid(row=0,column=9,sticky=EW)
        ttk.Label(self.routewin, text = "Select months:").grid(row=0,column=8,sticky=W)
        self.listboxR = Listbox(self.routewin, height=10, width=8, selectmode=EXTENDED, exportselection=FALSE)
        self.listboxR.grid(row=1,rowspan=5,column=8,sticky=EW)
        for t in self.periods:
            self.listboxR.insert(END,t)
        self.routewin.withdraw()

    def draw_CV(self):
        '''
        Draws GUI component: Pop-up window for editing C&V decisions
        '''

        self.user_add_cv = {}
        self.user_ex_cv = {}
        self.user_modality = {}
        self.cvwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.cvwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.cvwin))
        ttk.Label(self.cvwin, text = "C&V Procurement Decisions", font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky=W,pady=(0,5))
        ttk.Label(self.cvwin, text = "Include").grid(row=2,column=0,sticky=W)
        ttk.Label(self.cvwin, text = "Local Market", anchor=CENTER).grid(row=1,column=1)
        ttk.Label(self.cvwin, text = "Commodity", anchor=CENTER).grid(row=1,column=2)
        ttk.Label(self.cvwin, text = "Quantity (mt)", anchor=CENTER).grid(row=1,column=3)
        self.cv_add_src = StringVar()
        ttk.OptionMenu(self.cvwin, self.cv_add_src, "Select", *self.LMs).grid(row=2,column=1,sticky=EW)
        self.cv_add_src.trace('w',self.update_cv_add)
        self.cv_add_com = StringVar()
        self.cv_add_list = ttk.OptionMenu(self.cvwin, self.cv_add_com, "Select", *self.commodities)
        self.cv_add_list.grid(row=2,column=2,sticky=EW)
        self.cv_add_mt = StringVar()
        self.cv_add_mt.set("N/A")
        ttk.Entry(self.cvwin, textvariable = self.cv_add_mt, justify=CENTER).grid(row=2,column=3,sticky=EW)
        ttk.Button(self.cvwin, text = "Add constraint", command = self.include_cv).grid(row=2,column=4,sticky=EW)

        ttk.Label(self.cvwin, text = "Exclude").grid(row=3,column=0,sticky=W)
        self.cv_ex_src = StringVar()
        loclist = list(self.LMs)
        loclist.append("Any")
        ttk.OptionMenu(self.cvwin, self.cv_ex_src, "Select", *loclist).grid(row=3,column=1,sticky=EW)
        self.cv_ex_src.trace('w',self.update_cv_ex)
        self.cv_ex_com = StringVar()
        self.cv_ex_list = ttk.OptionMenu(self.cvwin, self.cv_ex_com, "Select", *self.commodities)
        self.cv_ex_list.grid(row=3,column=2,sticky=EW)
        ttk.Button(self.cvwin, text = "Add constraint", command = self.exclude_cv).grid(row=3,column=4,sticky=EW)
        r=4

        ttk.Label(self.cvwin, text = " ").grid(row=r,column=0,sticky=W)
        ttk.Label(self.cvwin, text = "Transfer Modality Decisions", font=("Helvetica",11,"bold")).grid(row=r+1,column=0,columnspan=3,sticky=W,pady=(0,5))
        ttk.Label(self.cvwin, text = "Min % C&V", anchor=CENTER).grid(row=r+2,column=2)
        ttk.Label(self.cvwin, text = "Max % C&V", anchor=CENTER).grid(row=r+2,column=3)
        ttk.Label(self.cvwin, text = "FDP ratio").grid(row=r+3,column=0,sticky=W)
        fdplist = list(self.FDPs)
        fdplist.append("All")
        self.cv_mod_fdp = StringVar()
        ttk.OptionMenu(self.cvwin, self.cv_mod_fdp, "Select", *fdplist).grid(row=r+3,column=1,sticky=EW)
        self.cv_mod_min = StringVar()
        self.cv_mod_min.set("0")
        ttk.Entry(self.cvwin, textvariable = self.cv_mod_min, justify=CENTER).grid(row=r+3,column=2,sticky=EW)
        self.cv_mod_max = StringVar()
        self.cv_mod_max.set("100")
        ttk.Entry(self.cvwin, textvariable = self.cv_mod_max, justify=CENTER).grid(row=r+3,column=3,sticky=EW)
        ttk.Button(self.cvwin, text = "Add constraint", command = self.set_mod).grid(row=r+3,column=4,sticky=EW)
        ttk.Label(self.cvwin, text = "Average ratio for CO").grid(row=r+4,column=0,columnspan=2,sticky=W)
        self.user_cv_min = StringVar()
        self.user_cv_min.set("0")
        self.user_cv_max = StringVar()
        self.user_cv_max.set("100")
        ttk.Entry(self.cvwin, textvariable = self.user_cv_min, justify=CENTER).grid(row=r+4, column=2, sticky=EW)
        ttk.Entry(self.cvwin, textvariable = self.user_cv_max, justify=CENTER).grid(row=r+4, column=3, sticky=EW)

        self.exp_pattern = {}
        self.exp_pattern["Cereals and Grains",0] = StringVar()
        self.exp_pattern["Cereals and Grains",0].set(20)
        self.exp_pattern["Cereals and Grains",1] = StringVar()
        self.exp_pattern["Cereals and Grains",1].set(40)
        self.exp_pattern["Vegetables and Fruits",0] = StringVar()
        self.exp_pattern["Vegetables and Fruits",0].set(15)
        self.exp_pattern["Vegetables and Fruits",1] = StringVar()
        self.exp_pattern["Vegetables and Fruits",1].set(30)
        self.exp_pattern["Other Food Items",0] = StringVar()
        self.exp_pattern["Other Food Items",0].set(0)
        self.exp_pattern["Other Food Items",1] = StringVar()
        self.exp_pattern["Other Food Items",1].set(25)
        self.exp_pattern["Non-Food Items",0] = StringVar()
        self.exp_pattern["Non-Food Items",0].set(10)
        self.exp_pattern["Non-Food Items",1] = StringVar()
        self.exp_pattern["Non-Food Items",1].set(50)
        r+=5
        ttk.Label(self.cvwin, text = " ").grid(row=r,column=0,sticky=W)
        l = ttk.Label(self.cvwin, text = "Beneficiary Expenditure Pattern", font=("Helvetica",11,"bold"))
        l.grid(row=r+1,column=0,columnspan=3,sticky=W,pady=(0,5))
        t = ("The Beneficiary Expenditure Pattern is only used when we opt for the 'CASH' approach to C&V."
            "\nUnder the 'CASH' approach, we make some assumptions on what beneficiaries will spend their"
            "\nnon-conditional transfer on in order to estimate the nutritional value of a cash contribution.")
        createToolTip(l,t)
        ttk.Label(self.cvwin, text = "Minimum %", anchor=CENTER).grid(row=r+2,column=1)
        ttk.Label(self.cvwin, text = "Maximum %", anchor=CENTER).grid(row=r+2,column=2)
        r+=3
        for g in ["Cereals and Grains","Vegetables and Fruits","Other Food Items","Non-Food Items"]:
            ttk.Label(self.cvwin, text = g).grid(row=r,column=0,sticky=W)
            ttk.Entry(self.cvwin, textvariable = self.exp_pattern[g,0], justify=CENTER).grid(row=r,column=1,sticky=EW)
            ttk.Entry(self.cvwin, textvariable = self.exp_pattern[g,1], justify=CENTER).grid(row=r,column=2,sticky=EW)
            r+=1

        ttk.Label(self.cvwin, text = "   ").grid(row=0,column=5,sticky=EW)
        ttk.Label(self.cvwin, text = "   ").grid(row=0,column=7,sticky=EW)
        ttk.Label(self.cvwin, text = "Select months:").grid(row=0,column=6,sticky=W)
        self.listboxCV = Listbox(self.cvwin, height=11,width=10, selectmode=EXTENDED, exportselection=FALSE)
        self.listboxCV.grid(row=1,rowspan=6,column=6,sticky=EW)
        for t in self.periods:
            self.listboxCV.insert(END,t)

        ttk.Label(self.cvwin, text = " ").grid(row=r,column=0,sticky=W)
        ttk.Button(self.cvwin, text = "Reset", command = self.reset_cv,width=10).grid(row=r+1,column=4,sticky=E)
        ttk.Button(self.cvwin, text = "Back", command = lambda: self.close(self.cvwin),width=10).grid(row=r+1,column=6,sticky=W)
        self.cvwin.withdraw()

    def draw_obj(self):
        '''
        Draws GUI component: Pop-up window for editing objectives decisions
        '''

        self.mingoal = {}
        self.maxgoal = {}
        self.objwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.objwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.objwin))
        ttk.Label(self.objwin, text = "Model Objectives", font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky=W,pady=(0,5))
        ttk.Label(self.objwin, text = "Target").grid(row=2,column=0,sticky=W)
        ttk.Label(self.objwin, text = "Statistic", anchor=CENTER).grid(row=1,column=1)
        ttk.Label(self.objwin, text = "Minimum", anchor=CENTER).grid(row=1,column=2)
        ttk.Label(self.objwin, text = "Maximum", anchor=CENTER).grid(row=1,column=3)
        ttk.Label(self.objwin, text = "Range:", anchor=CENTER).grid(row=1,column=4)
        self.statistic = StringVar()
        temp = list(self.stats.keys()) # The default order in Dictionaries is a complete mess
        temp.sort() # So we sort it alphabetically
        ttk.OptionMenu(self.objwin, self.statistic, "Total Costs", *temp).grid(row=2,column=1,sticky=EW)
        self.minstat = StringVar()
        self.minstat.set("N/A")
        ttk.Entry(self.objwin, textvariable = self.minstat, justify=CENTER).grid(row=2,column=2,sticky=EW)
        self.maxstat = StringVar()
        self.maxstat.set("N/A")
        ttk.Entry(self.objwin, textvariable = self.maxstat, justify=CENTER).grid(row=2,column=3,sticky=EW)
        self.statrange = StringVar()
        l1 = ["Selected months","Average","Total"]
        ttk.OptionMenu(self.objwin, self.statrange, l1[0], *l1).grid(row=2,column=4,sticky=EW)
        ttk.Button(self.objwin, text = "Add constraint", command = self.set_obj).grid(row=2,column=5,sticky=EW)

        ttk.Label(self.objwin, text = "   ").grid(row=0,column=6,sticky=EW)
        ttk.Label(self.objwin, text = "   ").grid(row=0,column=8,sticky=EW)
        ttk.Label(self.objwin, text = "Select months:").grid(row=0,column=7,sticky=W)
        self.listboxO = Listbox(self.objwin, height=7, width=10, selectmode=EXTENDED, exportselection=FALSE)
        self.listboxO.grid(row=1,rowspan=5,column=7,sticky=EW)
        for t in self.periods:
            self.listboxO.insert(END,t)

        ttk.Label(self.objwin, text = " ").grid(row=7,column=0,sticky=W)
        ttk.Button(self.objwin, text = "Back", command = lambda: self.close(self.objwin),width=10).grid(row=8,column=7,sticky=W)
        ttk.Button(self.objwin, text = "Reset", command = self.reset_obj,width=10).grid(row=8,column=5,sticky=E)
        self.objwin.withdraw()

    def draw_funding(self):
        '''
        Draws GUI component: Pop-up window for editing funding decisions
        '''

        self.user_add_ik = {}
        self.funwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.funwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.funwin))
        r = 0
        ttk.Label(self.funwin, text = "US In-Kind Donations", font = ("Helvetica",11,"bold")).grid(row=r,column=0,columnspan=3,sticky=W)
        ttk.Label(self.funwin, text = "Amount", anchor=CENTER).grid(row=r+1,column=0)
        ttk.Label(self.funwin, text = "Metric", anchor=CENTER).grid(row=r+1,column=1)
        ttk.Label(self.funwin, text = "Measurement", anchor=CENTER).grid(row=r+1,column=2)
        self.ik_donation = StringVar()
        self.ik_donation.set("N/A")
        ttk.Entry(self.funwin, textvariable = self.ik_donation, justify=CENTER).grid(row=r+2,column=0,sticky=EW)
        l1 = ["USD","MT"]
        self.ik_metric = StringVar()
        ttk.OptionMenu(self.funwin, self.ik_metric, l1[0], *l1).grid(row=r+2,column=1,sticky=EW)
        l2 = ["Value","Percentage"]
        self.ik_measure = StringVar()
        ttk.OptionMenu(self.funwin, self.ik_measure, l2[0], *l2).grid(row=r+2,column=2,sticky=EW)
        ttk.Button(self.funwin, text = "Add constraint", command = self.include_ik).grid(row=r+2,column=3,sticky=EW)
        r += 3

        ttk.Label(self.funwin, text = "   ").grid(row=0,column=5,sticky=EW)
        ttk.Label(self.funwin, text = "   ").grid(row=0,column=7,sticky=EW)
        ttk.Label(self.funwin, text = "Select months:").grid(row=0,column=6,sticky=W)
        self.listboxIK = Listbox(self.funwin, height=11,width=10, selectmode=EXTENDED, exportselection=FALSE)
        self.listboxIK.grid(row=1,rowspan=6,column=6,sticky=EW)
        for t in self.periods:
            self.listboxIK.insert(END,t)

        ttk.Label(self.funwin, text = " ").grid(row=10,column=0,sticky=EW)
        ttk.Button(self.funwin, text = "Reset", command = self.reset_ik,width=10).grid(row=10,column=3,sticky=E)
        ttk.Button(self.funwin, text = "Back", command = lambda: self.close(self.funwin),width=10).grid(row=10,column=6,sticky=W)
        self.funwin.withdraw()

    def draw_save(self):
        '''
        Draws GUI component: Save/load/display user input
        '''

        r,c = 0,5
        ttk.Label(self.frame_main, text = "Save/Load Input:", font = ("Helvetica",11,"bold"), anchor=CENTER).grid(row=r,column=c,columnspan=3,sticky=EW)
        ttk.Button(self.frame_main, text = "Reset Inputs", command = self.reset).grid(row=r+1,column=c,sticky=EW)
        ttk.Button(self.frame_main, text = "Display Inputs", command = self.display_inputs).grid(row=r+1,column=c+1,sticky=EW)
        b = ttk.Button(self.frame_main, text = "Quick Load", command = self.quick_load)
        b.grid(row=r+2,column=c,sticky=EW)
        createToolTip(b,"Return to the previously saved state")
        b = ttk.Button(self.frame_main, text = "Quick Save", command = self.quick_save)
        b.grid(row=r+3,column=c,sticky=EW)
        createToolTip(b,"Temporarily store the current user input\n(for permanent storage, use 'Save to .csv')")
        b = ttk.Button(self.frame_main, text = "Load From File", command = self.csv_load)
        b.grid(row=r+2,column=c+1,sticky=EW)
        createToolTip(b,"Load user input from a .csv file")
        b = ttk.Button(self.frame_main, text = "Save To File", command = self.csv_save)
        b.grid(row=r+3,column=c+1,sticky=EW)
        createToolTip(b,"Permanently store user input in a .csv file")
        b = ttk.Button(self.frame_main, text = "Convert Files", command = self.csv_convert)
        b.grid(row=r+1,column=c+2,sticky=EW)
        createToolTip(b,"Update all .csv files to newest format")
        self.csvnames = StringVar()
        self.csvnames.set("filename")
        ttk.Entry(self.frame_main, textvariable = self.csvnames,font=("Helvetica",8,"italic"),justify=CENTER).grid(row=r+3,column=c+2,sticky=EW)
        self.csvnamel = StringVar()
        self.savedscens = []
        self.csv_list = ttk.OptionMenu(self.frame_main, self.csvnamel, "Select", *self.savedscens)
        self.csv_list.grid(row=r+2,column=c+2,sticky=EW)

    def draw_analysis(self):
        '''
        Draws GUI component: Scenario analyses
        '''

        self.countscen = 1 # Keeps track of the scenarios
        self.prepped = ["","",""] # Used to check whether the constraints are prepared for the right time interval
        r,c = 5,5
        ttk.Label(self.frame_main, text = "Scenario Analysis:", font = ("Helvetica",11,"bold"), anchor=CENTER).grid(row=r,column=c,columnspan=3,sticky=EW)
        self.scenname = StringVar()
        self.scenname.set("Scenario_001")
        ttk.Entry(self.frame_main, textvariable = self.scenname,font=("Helvetica",8,"italic"),justify=CENTER).grid(row=r+1,column=c,columnspan=2,sticky=EW)
        b = ttk.Button(self.frame_main, text = "Solve Scenario", command = self.objset)
        b.grid(row=r+1,column=c+2,sticky=EW)
        createToolTip(b,"Find the optimal solution for the current input")
        b = ttk.Button(self.frame_main, text = "Scenario Analysis (From List)", command = lambda: self.show(self.scenwin))
        b.grid(row=r+2,column=c,columnspan=2,sticky=EW)
        createToolTip(b,"Solve multiple user-defined scenarios in a row")
        b = ttk.Button(self.frame_main, text = "Scenario Analysis (Auto)", command = lambda: self.show(self.autowin))
        b.grid(row=r+2,column=c+2,sticky=EW)
        createToolTip(b,"Access a range of automated analyses to gain some quick insights")

    def draw_outputs(self):
        '''
        Draws GUI component: Solution outputs/comparison
        '''

        r,c = 9,5
        ttk.Label(self.frame_main, text = "Solution Outputs:", font = ("Helvetica",11,"bold"), anchor=CENTER).grid(row=r,column=c,columnspan=3,sticky=EW)
        self.solutions = {} # will store benchmarks
        self.disp = {} # will store a subset of benchmarks for display
        script_dir = os.path.dirname(os.path.abspath(__file__))
        dest_dir = os.path.join(script_dir, 'output')
        ttk.Button(self.frame_main, text = "Display Benchmarks", command = self.display_benchmarks).grid(row=r+1,column=c,columnspan=2,sticky=EW)
        ttk.Button(self.frame_main, text = "Display Solution", command = lambda: self.show(self.solwin)).grid(row=r+2,column=c+2,sticky=EW)
        b = ttk.Button(self.frame_main, text = "Save Benchmarks To File", command = lambda: self.csv_benchmarks(dest_dir,"User"))
        b.grid(row=r+1,column=c+2,sticky=EW)
        createToolTip(b,"Save a range of scenario outputs to a .csv file\nTo be used to compare different scenarios")
        ttk.Button(self.frame_main, text = "Clear Benchmarks", command = self.clearbenchmarks).grid(row=r+2,column=c,columnspan=2,sticky=EW)

    def draw_tact(self):
        '''
        bla
        '''

        self.tactwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.tactwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.tactwin))
        ttk.Label(self.tactwin, text = "Filter Tactical Demand", font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky=W,pady=(0,5))

        l_fdp = list(self.tact_fdp.keys())
        l_fdp.sort()
        l_com = list(self.tact_com.keys())
        l_com.sort()
        l_mon = []
        for t in self.periods:
            if t in self.tact_mon.keys():
                l_mon.append(t)
        self.tactboxes = {}
        for i in (l_fdp + l_com + l_mon):
            self.tactboxes[i] = IntVar()
            self.tactboxes[i].set(1)

        ttk.Label(self.tactwin, text="FDP List").grid(row=1,column=0,sticky=EW)
        ttk.Label(self.tactwin, text="Metric Tonnes", anchor='e').grid(row=1,column=1,sticky=EW)
        r_fdp = 2
        for i in l_fdp:
            b = ttk.Checkbutton(self.tactwin, text = i, variable = self.tactboxes[i])
            b.grid(row=r_fdp,column=0,sticky=W)
            ttk.Label(self.tactwin, text = self.fmt_wcommas(self.tact_fdp[i])[1:-3], anchor='e').grid(row=r_fdp,column=1,sticky=EW)
            r_fdp += 1
        ttk.Label(self.tactwin, text = "   ").grid(row=0,column=2,sticky=EW)
        ttk.Label(self.tactwin, text="Commodity List").grid(row=1,column=3,sticky=EW)
        ttk.Label(self.tactwin, text="Metric Tonnes", anchor='e').grid(row=1,column=4,sticky=EW)
        r_com = 2
        for k in l_com:
            b = ttk.Checkbutton(self.tactwin, text = k, variable = self.tactboxes[k])
            b.grid(row=r_com,column=3,sticky=W)
            ttk.Label(self.tactwin, text = self.fmt_wcommas(self.tact_com[k])[1:-3], anchor='e').grid(row=r_com,column=4,sticky=EW)
            r_com += 1
        ttk.Label(self.tactwin, text = "   ").grid(row=0,column=5,sticky=EW)
        ttk.Label(self.tactwin, text="Month List").grid(row=1,column=6,sticky=EW)
        ttk.Label(self.tactwin, text="Metric Tonnes", anchor='e').grid(row=1,column=7,sticky=EW)
        r_mon = 2
        for t in l_mon:
            b = ttk.Checkbutton(self.tactwin, text = t, variable = self.tactboxes[t])
            b.grid(row=r_mon,column=6,sticky=W)
            ttk.Label(self.tactwin, text = self.fmt_wcommas(self.tact_mon[t])[1:-3], anchor='e').grid(row=r_mon,column=7,sticky=EW)
            r_mon += 1
        r = max(r_fdp, r_com, r_mon) + 1
        self.tactwin.grid_columnconfigure(0, weight=1, uniform="eq")
        self.tactwin.grid_columnconfigure(1, weight=1, uniform="eq")
        self.tactwin.grid_columnconfigure(3, weight=1, uniform="eq")
        self.tactwin.grid_columnconfigure(4, weight=1, uniform="eq")
        self.tactwin.grid_columnconfigure(6, weight=1, uniform="eq")
        self.tactwin.grid_columnconfigure(7, weight=1, uniform="eq")

        ttk.Label(self.tactwin, text = " ").grid(row=r,column=0,sticky=W)
        ttk.Button(self.tactwin, text = "Select All", command = lambda: [self.tactboxes[i].set(1) for i in l_fdp]).grid(row=r+1,column=0,sticky=EW)
        ttk.Button(self.tactwin, text = "Select None", command = lambda: [self.tactboxes[i].set(0) for i in l_fdp]).grid(row=r+1,column=1,sticky=EW)
        ttk.Button(self.tactwin, text = "Select All", command = lambda: [self.tactboxes[k].set(1) for k in l_com]).grid(row=r+1,column=3,sticky=EW)
        ttk.Button(self.tactwin, text = "Select None", command = lambda: [self.tactboxes[k].set(0) for k in l_com]).grid(row=r+1,column=4,sticky=EW)
        ttk.Button(self.tactwin, text = "Select All", command = lambda: [self.tactboxes[t].set(1) for t in l_mon]).grid(row=r+1,column=6,sticky=EW)
        ttk.Button(self.tactwin, text = "Select None", command = lambda: [self.tactboxes[t].set(0) for t in l_mon]).grid(row=r+1,column=7,sticky=EW)

        ttk.Button(self.tactwin, text = "Back", command = lambda: self.close(self.tactwin)).grid(row=r+2,column=9,sticky=EW)
        ttk.Label(self.tactwin, text = "   ").grid(row=0,column=8,sticky=EW)
        self.tactwin.withdraw()

    def draw_auto(self):
        '''
        Draws GUI component: Pop-up window with automated analyses
        '''

        self.nvs_scens = ["Supply 100% NVS","Supply   98% NVS","Supply   95% NVS","Supply   90% NVS","Supply   80% NVS"]
        self.cv_scens = ["Modality 100% C&V","Modality   75% C&V","Modality   50% C&V","Modality   25% C&V","Modality     0% C&V"]
        self.lt_scens = ["120 Day Response","  90 Day Response","  60 Day Response","  30 Day Response","  15 Day Response"]
        self.cur_scens = ["Remove 1 Commodity", "Replace 1 Commodity", "Optimise Ration Sizes", "Adjust Transfer Modality", "Increase Prices", "Scale Up Operation","Sourcing Breakdown","Adjust US-IK Funding","Allocate Resources"]
        self.scens = self.nvs_scens + self.cv_scens + self.lt_scens + self.cur_scens
        self.checkbox = {}
        for s in self.scens:
            self.checkbox[s] = IntVar()
            self.checkbox[s].set(0)

        self.autowin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.autowin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.autowin))
        ttk.Label(self.autowin, text = "Automated Scenario Analysis - Design", font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky=W,pady=(0,5))
        label_nvs = ttk.Label(self.autowin, text = "NVS Scenarios:", anchor=CENTER)
        label_nvs.grid(row=2,column=0)
        t = ("These scenarios explore the cost of nutrition"
            "\nGiven the user input, the model will try to find"
            "\nthe cheapest basket that supplies X% NVS")
        createToolTip(label_nvs,t)
        r=3
        for s in self.nvs_scens:
            b = ttk.Checkbutton(self.autowin, text = s, variable = self.checkbox[s])
            b.grid(row=r,column=0,sticky=W)
            createToolTip(b,t)
            r += 1
        label_cv = ttk.Label(self.autowin, text = "C&V Scenarios:", anchor=CENTER)
        label_cv.grid(row=2,column=1)
        t = ("These scenarios explore the cost of C&V"
            "\nGiven the user input, the model will try to find"
            "\nthe cheapest basket that supplies X% through C&V")
        createToolTip(label_cv,t)
        r=3
        for s in self.cv_scens:
            b = ttk.Checkbutton(self.autowin, text = s, variable = self.checkbox[s])
            b.grid(row=r,column=1,sticky=W)
            createToolTip(b,t)
            r += 1
        label_lt = ttk.Label(self.autowin, text = "LT Scenarios:", anchor=CENTER)
        label_lt.grid(row=2,column=2)
        t = ("These scenarios explore the cost of Lead Time"
            "\nGiven the user input, the model will try to find"
            "\nthe cheapest basket that can be supplied within X days")
        createToolTip(label_cv,t)
        r=3
        for s in self.lt_scens:
            b = ttk.Checkbutton(self.autowin, text = s, variable = self.checkbox[s])
            b.grid(row=r,column=2,sticky=W)
            createToolTip(b,t)
            r += 1

        ttk.Label(self.autowin, text = " ").grid(row=r,column=0,sticky=W)
        ttk.Label(self.autowin, text = "Automated Scenario Analysis - Trade-Off", font=("Helvetica",11,"bold")).grid(row=r+1, column=0,columnspan=3,sticky=W,pady=(0,5))
        ttk.Label(self.autowin, text = "Secondary Objective To Analyse").grid(row=r+2,columnspan=2,column=0,sticky=W)
        self.obj2 = StringVar()
        objplus = self.objectives
        objplus.append("None")
        ttk.OptionMenu(self.autowin, self.obj2, "None", *objplus).grid(row=r+2,column=2,sticky=EW)
        self.obj2.trace('w',self.update_obj)
        ttk.Label(self.autowin, text = "Minimum", anchor=CENTER).grid(row=r+3,column=0)
        ttk.Label(self.autowin, text = "Maximum", anchor=CENTER).grid(row=r+3,column=1)
        ttk.Label(self.autowin, text = "Increment", anchor=CENTER).grid(row=r+3,column=2)
        self.objmin = StringVar()
        self.objmin.set("N/A")
        ttk.Entry(self.autowin, textvariable = self.objmin, justify=CENTER).grid(row=r+4,column=0,sticky=EW)
        self.objmax = StringVar()
        self.objmax.set("N/A")
        ttk.Entry(self.autowin, textvariable = self.objmax, justify=CENTER).grid(row=r+4,column=1,sticky=EW)
        self.increment = StringVar()
        self.increment.set("N/A")
        ttk.Entry(self.autowin, textvariable = self.increment, justify=CENTER).grid(row=r+4,column=2,sticky=EW)
        r += 5
        ttk.Label(self.autowin, text = " ").grid(row=r,column=0)
        ttk.Label(self.autowin, text = "Sourcing trade-off per food group").grid(row=r+1,column=0,columnspan=2,sticky=W)
        groupoptions = list(self.foodgroups)
        groupoptions.append("None")
        groupoptions.append("All")
        self.breakdown = StringVar()
        ttk.OptionMenu(self.autowin,self.breakdown, "None", *groupoptions).grid(row=r+1,column=2,sticky=EW)
        r+=2

        ttk.Label(self.autowin, text = " ").grid(row=r,column=0,sticky=W)
        ttk.Label(self.autowin, text = "Automated Scenario Analysis - Adjust", font=("Helvetica",11,"bold")).grid(row=r+1,column=0,columnspan=3,sticky=W,pady=(0,5))
        script_dir = os.path.dirname(os.path.abspath(__file__))
        mypath = os.path.join(script_dir, 'saved')
        f = []
        for (dirpath, dirnames, filenames) in os.walk(mypath):
            f.extend(dirnames)
            break
        self.baseline = StringVar()
        if len(f)==0:
            self.baseline.set("No files found")
            f.append("NONE")
        elif ("Current basket") in f:
            self.baseline.set("Current basket")
        else:
            self.baseline.set(f[0])
        ttk.Label(self.autowin, text = "Baseline scenario").grid(row=r+2,column=0,sticky=W)
        self.auto_csv_list = ttk.OptionMenu(self.autowin, self.baseline, self.baseline.get(), *f)
        self.auto_csv_list.grid(row=r+2,column=1,columnspan=2,sticky=W)
        r=r+3
        tt = {}
        tt["Remove 1 Commodity"] = ("This analysis removes each commodity in the selected basket one-by-one"
                                    "\nThis will show the cost-effectiveness of each of the basket's components")
        tt["Replace 1 Commodity"] = ("This analysis replaces each commodity in the selected basket one-by-one"
                                    "\nwith a commodity from the same food group, using the same ration size"
                                    "\nThis will show the potential impact of alternative food basket compositions")
        tt["Optimise Ration Sizes"] = ("This analysis replaces the fixed ration sizes in the selected basket with"
                                    "\na ration 'interval', according to the specified percentage (X% deviation)"
                                    "\nThis will show the optimal ration sizes given the current food basket"
                                    "\ncomposition in the face of varying budget levels")
        tt["Adjust Transfer Modality"] = ("This analysis changes the percentage of metric tonnes that are procured through C&V"
                                    "\nThis will show the impact of different transfer modalities for the current food basket"
                                    "\non the performance of the supply chain")
        tt["Increase Prices"] =    ("This analysis changes Local and/or Regional procurement prices (-50% to +100%)"
                                    "\nThis will show how dependent the costs of the operation are on local prices"
                                    "\nand how the procurement ratio can help a CO adapt to price spikes")
        tt["Scale Up Operation"] = ("This analysis increases the amount of beneficiaries in each FDP by different percentages"
                                    "\nThis will show the cost of scaling up and the maximum size of the operation given"
                                    "\nthe current capacities")
        tt["Sourcing Breakdown"] = ("This analysis investigates all possible sourcing options for the current food basket"
                                    "\nThis will show the cost-competitiveness of all available sources, and allows users"
                                    "\nto quickly figure out which commodities to source through C&V or US-IK donations")
        tt["Adjust US-IK Funding"] = ("This analysis investigates different levels of US-IK funding (as percentage of total funding)"
                                    "\nThis will show which commodities are best sourced through US-IK funding and how this"
                                    "\nwill increase the cost of your operation as a whole.")
        tt["Allocate Resources"] = ("This analysis investigates how much each supplied activity contributes to the performance"
                                    "\nThis allows us to allocate the costs and capacities among projects and/or activities fairly"
                                    "\n  << UNDER CONSTRUCTION >>")
        self.deviation = StringVar()
        self.deviation.set("25%")
        ttk.Entry(self.autowin, textvariable = self.deviation, justify=CENTER).grid(row=r+2,column=1,sticky=EW)
        self.incr_cbt = IntVar()
        self.incr_cbt.set(0)
        ttk.Checkbutton(self.autowin, text = "CBT", variable = self.incr_cbt).grid(row=r+4,column=1,sticky=W)
        self.incr_loc = IntVar()
        self.incr_loc.set(0)
        ttk.Checkbutton(self.autowin, text = "Local", variable = self.incr_loc).grid(row=r+4,column=2,sticky=W)
        self.incr_reg = IntVar()
        self.incr_reg.set(0)
        ttk.Checkbutton(self.autowin, text = "Regional", variable = self.incr_reg).grid(row=r+4,column=3,sticky=W)
        for s in self.cur_scens:
            b = ttk.Checkbutton(self.autowin, text = s, variable = self.checkbox[s])
            b.grid(row=r,column=0,sticky=W)
            createToolTip(b,tt[s])
            r += 1
        self.remove = ""  # This var will store the commodity to be removed
        self.replace = [] # This var will store the commodities to be swapped
        self.ration = ""  # This var will store the ration size deviation
        self.totalmt = "" # This var will store the amount of mt to be bought
        self.mod_loc, self.mod_reg, self.mod_cbt = 1, 1, 1 # These vars capture price modifiers
        self.scaleup = 1  # This var will store the %-increase in demand

        self.checkreset = IntVar()
        self.checkreset.set(0)
        ttk.Label(self.autowin, text = " ").grid(row=r,column=0,sticky=W)
        ttk.Checkbutton(self.autowin, text = "Reset constraints before analysis", variable = self.checkreset).grid(row=r+1,column=0,columnspan=2,sticky=W)
        ttk.Button(self.autowin, text = "Select All", command = lambda: [self.checkbox[s].set(1) for s in self.scens]).grid(row=r,column=3,sticky=EW)
        ttk.Button(self.autowin, text = "Select None", command = lambda: [self.checkbox[s].set(0) for s in self.scens]).grid(row=r+1,column=3,sticky=EW)
        ttk.Button(self.autowin, text = "Analyse", command = self.autoanalysis).grid(row=r+2,column=0,sticky=EW)
        ttk.Button(self.autowin, text = "Back", command = lambda: self.close(self.autowin)).grid(row=r+2,column=3,sticky=EW)
        ttk.Label(self.autowin, text = "   ").grid(row=0,column=4,sticky=EW)
        self.autowin.withdraw()

    def draw_solution(self):
        '''
        Draws GUI component: Pop-up window for displaying variable values
        '''

        self.solwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.solwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.solwin))
        ttk.Label(self.solwin, text = "Last Solution's Outputs", font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=3,sticky=W,pady=(0,5))
        ttk.Button(self.solwin, text = "Food Basket", command = lambda: self.display_solution("Food Basket")).grid(row=2,column=0,sticky=EW)
        ttk.Button(self.solwin, text = "Nutrient Shortfalls", command = lambda: self.display_solution("Nutrient Shortfalls")).grid(row=3,column=0,sticky=EW)
        ttk.Button(self.solwin, text = "Sourcing Strategy", command = lambda: self.display_solution("Sourcing Strategy")).grid(row=4,column=0,sticky=EW)

        ttk.Label(self.solwin, text = "   ").grid(row=0,column=4,sticky=EW)
        ttk.Label(self.solwin, text = "   ").grid(row=0,column=6,sticky=EW)
        ttk.Label(self.solwin, text = "Select months:").grid(row=0,column=5,sticky=W)
        self.listboxSOL = Listbox(self.solwin, height=6, width=10, selectmode=EXTENDED, exportselection=FALSE)
        self.listboxSOL.grid(row=1,rowspan=5,column=5,sticky=EW)
        for t in self.periods:
            self.listboxSOL.insert(END,t)

        ttk.Label(self.solwin, text = " ").grid(row=6,column=0)
        ttk.Button(self.solwin, text = "Back", command = lambda: self.close(self.solwin)).grid(row=7,column=5,sticky=EW)
        self.solwin.withdraw()

    def draw_stdout(self):
        '''
        Draws GUI component: Output window (overrides console)
        '''

        l = ttk.Label(self.frame_output, text = "Tool Outputs:",font = ("Helvetica",11,"bold"))
        l.pack(side=TOP)
        self.text = Text(self.frame_output, wrap="word")
        self.text.pack(side=LEFT, fill=BOTH, expand=1)
        self.text.tag_configure("stderr", foreground="red")
        self.vsb = Scrollbar(self.frame_output, orient="vertical", command=self.text.yview)
        self.text.configure(yscrollcommand=self.vsb.set)
        self.vsb.pack(side=LEFT, fill=Y, expand=1)

        sys.stdout = TextRedirector(self.text, "stdout")
        sys.stderr = TextRedirector(self.text, "stderr")

    def draw_scenlist(self):
        '''
        Draws GUI component: Pop-up window for selecting scenarios to be analysed
        '''

        self.scenwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.scenwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.scenwin))
        self.scenwin.columnconfigure(0, weight=1)
        self.scenwin.columnconfigure(1, weight=1)
        self.scenwin.rowconfigure(1, weight=1)
        ttk.Label(self.scenwin, text = "Scenarios To Be Optimised:", font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=2,sticky=W,pady=(0,5))
        scrollbar = Scrollbar(self.scenwin, orient=VERTICAL)
        h = max(5, min(20,len(self.savedscens)))
        self.listbox_scen = Listbox(self.scenwin, height=h, selectmode=EXTENDED, yscrollcommand=scrollbar.set, exportselection=FALSE)
        self.listbox_scen.grid(row=1,column=0,columnspan=2,sticky=W+E+N+S)
        for s in self.savedscens:
            self.listbox_scen.insert(END,s)
        self.update_csv()
        scrollbar.config(command=self.listbox_scen.yview)
        scrollbar.grid(row=1,column=3,sticky=NS)

        ttk.Label(self.solwin, text = " ").grid(row=12,column=0)
        ttk.Button(self.scenwin, text = "Select All", command = lambda: [self.listbox_scen.selection_set(i) for i in range(0,self.listbox_scen.size())]).grid(row=13,column=0,sticky=EW)
        ttk.Button(self.scenwin, text = "Select None", command = lambda: [self.listbox_scen.selection_clear(i) for i in range(0,self.listbox_scen.size())]).grid(row=13,column=1,sticky=EW)
        ttk.Button(self.scenwin, text = "Analyse", command = self.listanalysis).grid(row=14,column=0,sticky=EW)
        ttk.Button(self.scenwin, text = "Back", command = lambda: self.close(self.scenwin)).grid(row=14,column=1,sticky=EW)
        self.scenwin.withdraw()

    def draw_activities(self):
        '''
        Draws GUI component: Pop-up window for selecting activities to be optimised
        '''

        self.actwin = Toplevel(background=self.bgcolor,padx=20,pady=20)
        self.actwin.protocol('WM_DELETE_WINDOW', lambda: self.close(self.actwin))
        self.actwin.columnconfigure(0, weight=1)
        self.actwin.columnconfigure(1, weight=1)
        self.actwin.rowconfigure(1, weight=1)
        self.activities = []
        self.old_ben = self.ben.get()
        ttk.Label(self.actwin, text = "Activities To Be Supplied:", font=("Helvetica",11,"bold")).grid(row=0,column=0,columnspan=2,sticky=W,pady=(0,5))
        scrollbar = Scrollbar(self.actwin, orient=VERTICAL)
        h = max(5, min(20,len(self.beneficiaries)))
        self.listbox_act = Listbox(self.actwin, height=h, selectmode=EXTENDED, yscrollcommand=scrollbar.set, exportselection=FALSE)
        self.listbox_act.grid(row=1,column=0,columnspan=2,sticky=W+E+N+S)
        for b in self.beneficiaries:
            if b == self.ben.get():
                continue # We only consider the activities that we are not optimising
            self.listbox_act.insert(END,b)
        scrollbar.config(command=self.listbox_act.yview)
        scrollbar.grid(row=1,column=3,sticky=NS)

        ttk.Label(self.solwin, text = " ").grid(row=12,column=0)
        ttk.Button(self.actwin, text = "Select All", command = lambda: [self.listbox_act.selection_set(i) for i in range(0,self.listbox_act.size())]).grid(row=13,column=0,sticky=EW)
        ttk.Button(self.actwin, text = "Select None", command = lambda: [self.listbox_act.selection_clear(i) for i in range(0,self.listbox_act.size())]).grid(row=13,column=1,sticky=EW)
        ttk.Button(self.actwin, text = "Set Activities", command = lambda: self.set_act()).grid(row=14,column=0,sticky=EW)
        ttk.Button(self.actwin, text = "Back", command = lambda: self.close(self.actwin)).grid(row=14,column=1,sticky=EW)
        self.actwin.withdraw()

    def reset_act(self):
        '''
        Reset user constraints: Activities
        '''

        self.ben.set(self.benlist[0])
        self.activities = []
        for b in self.beneficiaries:
            if b == self.ben.get():
                continue # We only consider the activities that we are not optimising
            self.listbox_act.insert(END,b)
        for i in range(self.listbox_act.size()):
            self.listbox_act.itemconfig(i, background="white")
        self.act_button.configure(text="Select (0)")
        self.old_ben = self.ben.get()

    def reset_fix(self):
        '''
        Reset user constraints: Fixed commodities
        '''

        self.food2fix = []
        for i in range(0,15):
            self.fix_spec[i].set("Filter")
            self.fix_com[i].set("Select")
            self.fix_quant[i].set("N/A")

    def reset_route(self):
        '''
        Reset user constraints: Routing
        '''

        self.user_add_route = {}
        self.user_ex_route = {}
        self.user_cap_util = {}
        self.user_cap_aloc = {}
        self.route_util_min.set("0")
        self.route_util_max.set("100")
        self.route_aloc_min.set("0")
        self.route_aloc_max.set("100")

    def reset_proc(self):
        '''
        Reset user constraints: Procurement
        '''

        self.user_add_proc_int = {}
        self.user_add_proc_loc = {}
        self.user_ex_proc_int = {}
        self.user_ex_proc_loc = {}
        self.user_int_min.set("0")
        self.user_reg_min.set("0")
        self.user_loc_min.set("0")
        self.user_int_max.set("100")
        self.user_reg_max.set("100")
        self.user_loc_max.set("100")

    def reset_fb(self):
        '''
        Reset user constraints: Food basket
        '''

        for k in self.commodities:
            self.user_add_com[k]=[0,1000]
        for k in self.supcom:
            if k in self.user_add_com.keys():
                self.user_add_com.pop(k,None)
        self.user_ex_com = []
        self.user_add_nut = {}
        self.user_add_fg = {}
        self.user_ex_fg = []
        self.user_add_mincom.set("N/A")
        self.user_add_maxcom.set("N/A")
        self.user_nut_minprot.set("0")
        self.user_nut_maxprot.set("100")
        self.user_nut_minfat.set("0")
        self.user_nut_maxfat.set("100")
        self.gmo.set(1)

    def reset_cv(self):
        '''
        Reset user constraints: C&V
        '''

        self.user_add_cv = {}
        self.user_ex_cv = {}
        self.user_modality = {}
        self.user_cv_min.set("0")
        self.user_cv_max.set("100")
        self.cv_add_src.set("Select")
        self.cv_add_com.set("Select")
        self.cv_add_mt.set("N/A")
        self.cv_ex_src.set("Select")
        self.cv_ex_com.set("Select")
        self.cv_mod_fdp.set("Select")
        self.cv_mod_min.set("0")
        self.cv_mod_max.set("100")
        self.exp_pattern["Cereals and Grains",0].set(20)
        self.exp_pattern["Cereals and Grains",1].set(40)
        self.exp_pattern["Vegetables and Fruits",0].set(15)
        self.exp_pattern["Vegetables and Fruits",1].set(30)
        self.exp_pattern["Other Food Items",0].set(0)
        self.exp_pattern["Other Food Items",1].set(25)
        self.exp_pattern["Non-Food Items",0].set(10)
        self.exp_pattern["Non-Food Items",1].set(50)

    def reset_ik(self):
        '''
        Reset user constraints: In-Kind donations
        '''

        self.user_add_ik = {}
        self.ik_donation.set("N/A")

    def reset_obj(self):
        '''
        Reset user constraints: Objectives
        '''

        self.mingoal = {}
        self.minstat.set("N/A")
        self.maxstat.set("N/A")
        self.statrange.set("Selected months")
        self.maxgoal = {}

    def reset(self):
        '''
        Reset user constraints: All
        '''

        self.reset_route()
        self.reset_proc()
        self.reset_fb()
        self.reset_cv()
        self.reset_fix()
        self.reset_obj()
        self.reset_ik()
        # self.reset_act()
        self.allowshortfalls.set(0)
        self.sensible.set(1)
        self.useforecasts.set(1)
        self.supply_tact.set(0)
        for i in self.tactboxes.keys():
            self.tactboxes[i].set(1)
        self.varbasket.set("Variable")
        self.modality.set("Voucher")
        print "-- Constraints Reset --"
        print " "

    def clearbenchmarks(self):
        '''
        Clear stored solutions (cross-comparison)
        Output files (.csv) will still exist
        '''

        self.solutions={}
        self.disp={}
        self.countscen=1
        self.scenname.set("Scenario_"+str(self.countscen).zfill(3))
        print "-- Benchmarks Cleared --"
        print " "

    def show(self,window):
        '''
        GUI: Show the window
        '''

        window.deiconify()

    def close(self,window):
        '''
        GUI: Hide the window
        '''

        window.withdraw()

    def update_conversion(self, *args):
        '''
        Update: Conversion table (in fixed food basket window)
        '''

        check = 1
        try:
            d, s, c, x = self.fix_days.get(), self.fix_hh.get(), self.fix_g2l.get(), self.fix_in.get()
            u1, u2 = self.fix_unit1.get(), self.fix_unit2.get()
        except ValueError:
            check = 0

        if check == 1:
            for i in xrange(1, len(self.fix_def.keys())/2 + 1):
                p = self.fix_def[i,1].cget("text")
                v = float(p[:-3])
                if p.endswith("KG"):
                    c = self.convert_unit(v, "KG/hh/m", "g/p/d")
                    self.fix_def[i,2].config(text=str(c))
                else: # endswith "L"
                    c = self.convert_unit(v, "L/hh/m", "g/p/d")
                    self.fix_def[i,2].config(text=str(c))
            self.fix_out.config(text=str(self.convert_unit(x,u1,u2)))

    def update_act(self, *args):
        '''
        Update: Activities list (in draw_activities(self))
        '''

        for i in range(self.listbox_act.size()):
            if self.listbox_act.get(i) == self.ben.get():
                # Update self.listbox_act
                self.listbox_act.delete(i)
                self.listbox_act.insert(i,self.old_ben)
                # Update self.activities if necessary
                if self.ben.get() in self.activities:
                    self.activities.remove(self.ben.get())
                    self.activities.append(self.old_ben) # Swap self.ben.get() with the previously optimised activity
                    self.listbox_act.itemconfig(i, background="light sky blue")
                break
        self.old_ben = self.ben.get()
        self.fix_days.set(max(self.feedingdays[self.ben.get(),k] for k in self.commodities))

    def update_csv(self):
        '''
        Update: Saved scenario list
        '''

        script_dir = os.path.dirname(os.path.abspath(__file__))
        mypath = os.path.join(script_dir, 'saved')
        self.savedscens = []
        for (dirpath, dirnames, filenames) in os.walk(mypath):
            self.savedscens.extend(dirnames)
            break
        for f in self.savedscens:
            if f.startswith('backup') or f.startswith('Backup'):
                self.savedscens.remove(f)

        if len(self.savedscens) > 0:
            menu = self.csv_list['menu']
            menu2 = self.auto_csv_list['menu']
            menu.delete(0, 'end')
            menu2.delete(0, 'end')
            if self.csvnamel.get()=="Select":
                if "Current basket" in self.savedscens:
                    self.csvnamel.set("Current basket")
                else:
                    self.csvnamel.set(self.savedscens[0])
            for k in self.savedscens:
                menu.add_command(label=k, command=lambda k=k: self.csvnamel.set(k))
                menu2.add_command(label=k, command=lambda k=k: self.baseline.set(k))
            self.auto_csv_list

            # update self.listbox_scen
            self.listbox_scen.delete(0,END)
            for s in self.savedscens:
                self.listbox_scen.insert(END,s)
            h = max(5, min(20,len(self.savedscens)))
            self.listbox_scen.configure(height=h)


    def update_lists(self, *args):
        '''
        This function is called when the user changes the time horizon.
        It updates all listboxes that are based on the time horizon to reflect this change.
        '''

        time = list(self.periods)
        for t in self.periods:
            if t != self.tstart.get():
                time.remove(t)
            else:
                break
        for t in reversed(self.periods):
            if t != self.tend.get():
                time.remove(t)
            else:
                break

        temp = [self.listboxR,self.listboxP,self.listboxO,self.listboxCV,self.listboxSOL,self.listboxIK]
        for box in temp:
            box.delete(0,END)
            for t in time:
                box.insert(END,t)

    def update_fix_com(self, *args):
        '''
        This function updates the list for fix_com, based on the user's choice for fix_spec
        '''

        try:
            index = int(args[0][-2:])
        except:
            index = int(args[0][-1:])
        i = (index-self.filter)/3
        # NB: this is an extreeeeeeeeemely roundabout way to link to the appropriate menu, but it works
        menu = self.fix_list[i]['menu']
        menu.delete(0, 'end')
        c = self.fix_spec[i].get()
        if c != "Filter":
            for k in self.commodities:
                if self.sup[k] == c:
                    menu.add_command(label=k, command=lambda k=k: self.fix_com[i].set(k))
        else:
            for k in self.commodities:
                menu.add_command(label=k, command=lambda k=k: self.fix_com[i].set(k))

    def update_fb_add_speccom(self, *args):
        '''
        This function updates the list for fb_add_speccom, based on the user's choice for fb_add_com
        '''

        menu = self.fb_add_list['menu']
        menu.delete(0, 'end')
        if self.fb_add_com.get() != "Filter":
            st = "Any"
            self.fb_add_speccom.set(st)
            menu.add_command(label=st, command=lambda st=st: self.fb_add_speccom.set(st))
            for k in self.commodities:
                if self.sup[k] == self.fb_add_com.get():
                    menu.add_command(label=k, command=lambda k=k: self.fb_add_speccom.set(k))
        else:
            st = "Select"
            self.fb_add_speccom.set(st)
            menu.add_command(label=st, command=lambda st=st: self.fb_add_speccom.set(st))
            for k in self.commodities:
                menu.add_command(label=k, command=lambda k=k: self.fb_add_speccom.set(k))

    def update_fb_ex_speccom(self, *args):
        '''
        This function updates the list for fb_ex_speccom, based on the user's choice for fb_ex_com
        '''

        menu = self.fb_ex_list['menu']
        menu.delete(0, 'end')
        if self.fb_ex_com.get() != "Filter":
            st = "Any"
            self.fb_ex_speccom.set(st)
            menu.add_command(label=st, command=lambda st=st: self.fb_ex_speccom.set(st))
            for k in self.commodities:
                if self.sup[k] == self.fb_ex_com.get():
                    menu.add_command(label=k, command=lambda k=k: self.fb_ex_speccom.set(k))
        else:
            st = "Select"
            self.fb_ex_speccom.set(st)
            menu.add_command(label=st, command=lambda st=st: self.fb_ex_speccom.set(st))
            for k in self.commodities:
                menu.add_command(label=k, command=lambda k=k: self.fb_ex_speccom.set(k))

    def update_proc_add_ndp_int(self, *args):
        '''
        This function updates the list for proc_add_ndp_int,
        based on the user's choice for proc_add_src_int and proc_add_inco_int
        '''

        menu = self.proc_add_list_ndp_int['menu']
        menu.delete(0, 'end')
        c = self.proc_add_src_int.get()
        i = self.proc_add_inco_int.get()
        if i != "Select":
            st = "Any"
            self.proc_add_ndp_int.set(st)
            menu.add_command(label=st, command=lambda st=st: self.proc_add_ndp_int.set(st))
            ndplist = []
            if c == "Any":
                if i == "Any":
                    for arc in self.proccap.keys(): # arc = (origin, destination, com)
                        if arc[1] in (self.ISs+self.RSs):
                            ndplist.append(arc[1])
                else:
                    for arc in self.proccap.keys():
                        if arc[1] in (self.ISs+self.RSs) and arc[0].endswith(i):
                            ndplist.append(arc[1])
            else:
                if i == "Any":
                    for arc in self.proccap.keys():
                        if arc[1] in (self.ISs+self.RSs) and arc[0].startswith(c):
                            ndplist.append(arc[1])
                else:
                    for arc in self.proccap.keys():
                        if arc[1] in (self.ISs+self.RSs) and arc[0] == c + " - " + i:
                            ndplist.append(arc[1])

            ndplist = list(set(ndplist))
            ndplist.sort()
            for k in ndplist:
                menu.add_command(label=k, command=lambda k=k: self.proc_add_ndp_int.set(k))
            if len(ndplist)==0:
                self.proc_add_ndp_int.set("NOT FOUND")
        else:
            self.proc_add_ndp_int.set("Select")

    def update_proc_add_inco_int(self, *args):
        '''
        This function sets the incoterm to "Any"
        '''

        self.proc_add_inco_int.set("Any")

    def update_proc_add_com_int(self, *args):
        '''
        This function updates the list for proc_add_com_int, based on the user's choice for
        proc_add_src_int, proc_add_inco_int, and proc_add_ndp_int
        '''

        menu = self.proc_add_list_com_int['menu']
        menu.delete(0,'end')
        self.proc_add_com_int.set("Select")
        c = self.proc_add_src_int.get()
        i = self.proc_add_inco_int.get()
        l = self.proc_add_ndp_int.get()
        if l != "Select":
            templist = []
            if c == "Any":
                if i == "Any":
                    if l == "Any":
                        for proc in self.proccap.keys(): # proc = (src, ndp, com)
                            if proc[1] in (self.ISs+self.RSs):
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l:
                                templist.append(proc[2])
                else:
                    if l == "Any":
                        for proc in self.proccap.keys():
                            if proc[1] in (self.ISs+self.RSs) and proc[0].endswith(i):
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l and proc[0].endswith(i):
                                templist.append(proc[2])
            else:
                if i == "Any":
                    if l == "Any":
                        for proc in self.proccap.keys(): # proc = (src, ndp, com)
                            if proc[1] in (self.ISs+self.RSs) and proc[0].startswith(c):
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l and proc[0].startswith(c):
                                templist.append(proc[2])
                else:
                    if l == "Any":
                        for proc in self.proccap.keys():
                            if proc[1] in (self.ISs+self.RSs) and proc[0]==c + " - " + i:
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l and proc[0]==c + " - " + i:
                                templist.append(proc[2])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.proc_add_com_int.set(k))
            if len(templist)==0:
                self.proc_add_com_int.set("NOT FOUND")
        else:
            self.proc_add_com_int.set("Select")

    def update_proc_add_ndp_loc(self, *args):
        '''
        This function updates the list for proc_add_ndp_loc, based on the user's choice for
        proc_add_src_loc and proc_add_inco_loc
        '''

        menu = self.proc_add_list_ndp_loc['menu']
        menu.delete(0, 'end')
        c = self.proc_add_src_loc.get()
        i = self.proc_add_inco_loc.get()
        if i != "Select":
            st = "Any"
            self.proc_add_ndp_loc.set(st)
            menu.add_command(label=st, command=lambda st=st: self.proc_add_ndp_loc.set(st))
            ndplist = []
            if c == "Any":
                if i == "Any":
                    for arc in self.proccap.keys(): # arc = (origin, destination, com)
                        if arc[1] in self.LSs:
                            ndplist.append(arc[1])
                else:
                    for arc in self.proccap.keys():
                        if arc[1] in self.LSs and arc[0].endswith(i):
                            ndplist.append(arc[1])
            else:
                if i == "Any":
                    for arc in self.proccap.keys():
                        if arc[1] in self.LSs and arc[0].startswith(c):
                            ndplist.append(arc[1])
                else:
                    for arc in self.proccap.keys():
                        if arc[1] in self.LSs and arc[0] == c + " - " + i:
                            ndplist.append(arc[1])

            ndplist = list(set(ndplist))
            ndplist.sort()
            for k in ndplist:
                menu.add_command(label=k, command=lambda k=k: self.proc_add_ndp_loc.set(k))
            if len(ndplist)==0:
                self.proc_add_ndp_loc.set("NOT FOUND")
        else:
            self.proc_add_ndp_loc.set("Select")

    def update_proc_add_inco_loc(self, *args):
        '''
        This function sets the Incoterm to "Any"
        '''

        self.proc_add_inco_loc.set("Any")

    def update_proc_add_com_loc(self, *args):
        '''
        This function updates the list for proc_add_com_loc, based on the user's choice for
        proc_add_src_loc, proc_add_inco_loc, and proc_add_ndp_loc
        '''

        menu = self.proc_add_list_com_loc['menu']
        menu.delete(0,'end')
        self.proc_add_com_loc.set("Select")
        c = self.proc_add_src_loc.get()
        i = self.proc_add_inco_loc.get()
        l = self.proc_add_ndp_loc.get()
        if l != "Select":
            templist = []
            if c == "Any":
                if i == "Any":
                    if l == "Any":
                        for proc in self.proccap.keys(): # proc = (src, ndp, com)
                            if proc[1] in self.LSs:
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l:
                                templist.append(proc[2])
                else:
                    if l == "Any":
                        for proc in self.proccap.keys():
                            if proc[1] in self.LSs and proc[0].endswith(i):
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l and proc[0].endswith(i):
                                templist.append(proc[2])
            else:
                if i == "Any":
                    if l == "Any":
                        for proc in self.proccap.keys(): # proc = (src, ndp, com)
                            if proc[1] in self.LSs and proc[0].startswith(c):
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l and proc[0].startswith(c):
                                templist.append(proc[2])
                else:
                    if l == "Any":
                        for proc in self.proccap.keys():
                            if proc[1] in self.LSs and proc[0]==c + " - " + i:
                                templist.append(proc[2])
                    else:
                        for proc in self.proccap.keys():
                            if proc[1] == l and proc[0]==c + " - " + i:
                                templist.append(proc[2])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.proc_add_com_loc.set(k))
            if len(templist)==0:
                self.proc_add_com_loc.set("NOT FOUND")
        else:
            self.proc_add_com_loc.set("Select")

    def update_proc_ex_ndp_int(self, *args):
        '''
        This function updates the list for proc_ex_ndp_int, based on the user's choice for proc_ex_src_int
        '''

        menu = self.proc_ex_list_ndp_int['menu']
        menu.delete(0,'end')
        c = self.proc_ex_src_int.get()
        if c != "Select":
            st = "Any"
            self.proc_ex_ndp_int.set(st)
            menu.add_command(label=st, command=lambda st=st: self.proc_ex_ndp_int.set(st))
            templist = []
            if c == "Any":
                for proc in self.proccap.keys():
                    if proc[1] in (self.ISs+self.RSs):
                        templist.append(proc[1])
            else:
                for proc in self.proccap.keys():
                    if proc[0].startswith(c) and proc[1] in (self.ISs+self.RSs):
                        templist.append(proc[1])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.proc_ex_ndp_int.set(k))
            if len(templist)==0:
                self.proc_ex_ndp_int.set("NOT FOUND")
        else:
            self.proc_ex_ndp_int.set("Select")

    def update_proc_ex_com_int(self, *args):
        '''
        This function updates the list for proc_ex_com_int, based on the user's choice for
        proc_ex_src_int and proc_ex_ndp_int
        '''

        menu = self.proc_ex_list_com_int['menu']
        menu.delete(0,'end')
        c = self.proc_ex_src_int.get()
        l = self.proc_ex_ndp_int.get()
        if l != "Select":
            st = "Any"
            self.proc_ex_com_int.set(st)
            menu.add_command(label=st, command=lambda st=st: self.proc_ex_com_int.set(st))
            templist = []
            if c == "Any":
                if l == "Any":
                    for proc in self.proccap.keys():
                        if proc[1] in (self.ISs+self.RSs):
                            templist.append(proc[2])
                else:
                    for proc in self.proccap.keys():
                        if proc[1] == l:
                            templist.append(proc[2])
            else:
                if l == "Any":
                    for proc in self.proccap.keys():
                        if proc[0].startswith(c) and proc[1] in (self.ISs+self.RSs):
                            templist.append(proc[2])
                else:
                    for proc in self.proccap.keys():
                        if proc[0].startswith(c) and proc[1] == l:
                            templist.append(proc[2])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.proc_ex_com_int.set(k))
            if len(templist)==0:
                self.proc_ex_com_int.set("NOT FOUND")
        else:
            self.proc_ex_com_int.set("Select")

    def update_proc_ex_ndp_loc(self, *args):
        '''
        This function updates the list for proc_ex_ndp_loc, based on the user's choice for proc_ex_src_loc
        '''

        menu = self.proc_ex_list_ndp_loc['menu']
        menu.delete(0,'end')
        c = self.proc_ex_src_loc.get()
        if c != "Select":
            st = "Any"
            self.proc_ex_ndp_loc.set(st)
            menu.add_command(label=st, command=lambda st=st: self.proc_ex_ndp_loc.set(st))
            templist = []
            if c == "Any":
                for proc in self.proccap.keys():
                    if proc[1] in self.LSs:
                        templist.append(proc[1])
            else:
                for proc in self.proccap.keys():
                    if proc[0].startswith(c) and proc[1] in self.LSs:
                        templist.append(proc[1])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.proc_ex_ndp_loc.set(k))
            if len(templist)==0:
                self.proc_ex_ndp_loc.set("NOT FOUND")
        else:
            self.proc_ex_ndp_loc.set("Select")

    def update_proc_ex_com_loc(self, *args):
        '''
        This function updates the list for proc_ex_com_loc, based on the user's choice for
        proc_ex_src_loc and proc_ex_ndp_loc
        '''

        menu = self.proc_ex_list_com_loc['menu']
        menu.delete(0,'end')
        c = self.proc_ex_src_loc.get()
        l = self.proc_ex_ndp_loc.get()
        if l != "Select":
            st = "Any"
            self.proc_ex_com_loc.set(st)
            menu.add_command(label=st, command=lambda st=st: self.proc_ex_com_loc.set(st))
            templist = []
            if c == "Any":
                if l == "Any":
                    for proc in self.proccap.keys():
                        if proc[1] in self.LSs:
                            templist.append(proc[2])
                else:
                    for proc in self.proccap.keys():
                        if proc[1] == l:
                            templist.append(proc[2])
            else:
                if l == "Any":
                    for proc in self.proccap.keys():
                        if proc[0].startswith(c) and proc[1] in self.LSs:
                            templist.append(proc[2])
                else:
                    for proc in self.proccap.keys():
                        if proc[0].startswith(c) and proc[1] == l:
                            templist.append(proc[2])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.proc_ex_com_loc.set(k))
            if len(templist)==0:
                self.proc_ex_com_loc.set("NOT FOUND")
        else:
            self.proc_ex_com_loc.set("Select")

    def update_route_add_loc1(self, *args):
        '''
        This function updates the list for route_add_loc1, based on the user's choice for route_add_type
        '''

        menu = self.route_add_list_loc1['menu']
        menu.delete(0,'end')
        if self.route_add_type.get() != "Select":
            self.route_add_loc1.set("Select")
            self.route_add_com.set("Any")
            if self.route_add_type.get() == "Load Port":
                for k in self.ISs:
                    menu.add_command(label=k, command=lambda k=k: self.route_add_loc1.set(k))
            elif self.route_add_type.get() == "Discharge Port":
                for k in self.DPs:
                    menu.add_command(label=k, command=lambda k=k: self.route_add_loc1.set(k))
            elif self.route_add_type.get() == "Extended Distribution Point":
                for k in self.EDPs:
                    menu.add_command(label=k, command=lambda k=k: self.route_add_loc1.set(k))
            elif self.route_add_type.get() == "Regional Market":
                for k in self.RSs:
                    menu.add_command(label=k, command=lambda k=k: self.route_add_loc1.set(k))
            elif self.route_add_type.get() == "Local Supplier":
                for k in self.LSs:
                    menu.add_command(label=k, command=lambda k=k: self.route_add_loc1.set(k))
            else: # loctype = Local Market
                for k in self.LMs:
                    menu.add_command(label=k, command=lambda k=k: self.route_add_loc1.set(k))
        else:
            for k in (self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs):
                menu.add_command(label=k, command=lambda k=k: self.route_add_loc1.set(k))

    def update_route_add_loc2(self, *args):
        '''
        This function updates the list for route_add_loc2, based on the user's choice for route_add_loc1
        '''

        menu = self.route_add_list_loc2['menu']
        menu.delete(0,'end')
        if self.route_add_loc1.get() != "Select":
            st = "Any"
            self.route_add_loc2.set(st)
            self.route_add_com.set(st)
            menu.add_command(label=st, command=lambda st=st: self.route_add_loc2.set(st))
            templist = []
            for key in self.cost.keys():
                if key[0]==self.route_add_loc1.get():
                    templist.append(key[1])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.route_add_loc2.set(k))
            if len(templist)==0:
                self.route_add_loc2.set("NOT FOUND")
        else:
            self.route_add_loc2.set("Select")
            for k in (self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs):
                menu.add_command(label=k, command=lambda k=k: self.route_add_loc2.set(k))

    def update_route_ex_loc1(self, *args):
        '''
        This function updates the list for route_ex_loc1, based on the user's choice for route_ex_type
        '''

        menu = self.route_ex_list_loc1['menu']
        menu.delete(0,'end')
        if self.route_ex_type.get() != "Select":
            self.route_ex_loc1.set("Select")
            self.route_ex_loc2.set("Select")
            self.route_ex_com.set("Any")
            if self.route_ex_type.get() == "Load Port":
                for k in self.ISs:
                    menu.add_command(label=k, command=lambda k=k: self.route_ex_loc1.set(k))
            elif self.route_ex_type.get() == "Discharge Port":
                for k in self.DPs:
                    menu.add_command(label=k, command=lambda k=k: self.route_ex_loc1.set(k))
            elif self.route_ex_type.get() == "Extended Distribution Point":
                for k in self.EDPs:
                    menu.add_command(label=k, command=lambda k=k: self.route_ex_loc1.set(k))
            elif self.route_ex_type.get() == "Regional Market":
                for k in self.RSs:
                    menu.add_command(label=k, command=lambda k=k: self.route_ex_loc1.set(k))
            elif self.route_ex_type.get() == "Local Supplier":
                for k in self.LSs:
                    menu.add_command(label=k, command=lambda k=k: self.route_ex_loc1.set(k))
            else: # loctypeb = Local Market
                for k in self.LMs:
                    menu.add_command(label=k, command=lambda k=k: self.route_ex_loc1.set(k))
        else:
            for k in (self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs):
                menu.add_command(label=k, command=lambda k=k: self.route_ex_loc1.set(k))

    def update_route_ex_loc2(self, *args):
        '''
        This function updates the list for route_ex_loc2, based on the user's choice for route_ex_loc1
        '''

        menu = self.route_ex_list_loc2['menu']
        menu.delete(0,'end')
        if self.route_ex_loc1.get() != "Select":
            st = "Any"
            self.route_ex_com.set(st)
            self.route_ex_loc2.set(st)
            menu.add_command(label=st, command=lambda st=st: self.route_ex_loc2.set(st))
            templist = []
            for key in self.cost.keys():
                if key[0]==self.route_ex_loc1.get():
                    templist.append(key[1])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.route_ex_loc2.set(k))
            if len(templist)==0:
                self.route_ex_loc2.set("NOT FOUND")
        else:
            self.route_ex_loc2.set("Select")
            for k in (self.ISs+self.DPs+self.EDPs+self.RSs+self.LMs):
                menu.add_command(label=k, command=lambda k=k: self.route_ex_loc2.set(k))

    def update_obj(self, *args):
        '''
        This function fills in the base min/max values for trade-off analyses
        '''

        s = self.obj2.get()
        if s != "None":
            self.objmin.set(self.base[s][0])
            self.objmax.set(self.base[s][1])
            self.increment.set(self.base[s][2])
        else:
            self.objmin.set("N/A")
            self.objmax.set("N/A")
            self.increment.set("N/A")

    def update_cv_ex(self, *args):
        '''
        This function updates the list for cv_ex_com, based on the user's choice for cv_ex_src
        '''

        menu = self.cv_ex_list['menu']
        menu.delete(0,'end')
        if self.cv_ex_src.get() != "Select":
            templist = []
            for key in self.proccap.keys():
                if self.cv_ex_src.get() == "Any" and key[1] in self.LMs:
                    templist.append(key[2])
                elif key[1]==self.cv_ex_src.get():
                    templist.append(key[2])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.cv_ex_com.set(k))
            if len(templist)==0:
                self.cv_ex_com.set("NOT FOUND")
            else:
                k = "Any"
                self.cv_ex_com.set(k)
                menu.add_command(label=k, command=lambda k=k: self.cv_ex_com.set(k))
        else:
            self.cv_ex_com.set("Select")
            for k in self.commodities:
                menu.add_command(label=k, command=lambda k=k: self.cv_ex_com.set(k))
            k = "Any"
            menu.add_command(label=k, command=lambda k=k: self.cv_ex_com.set(k))

    def update_cv_add(self, *args):
        '''
        This function updates the list for cv_add_com, based on the user's choice for cv_add_src
        '''

        menu = self.cv_add_list['menu']
        menu.delete(0,'end')
        if self.cv_add_src.get() != "Select":
            self.cv_add_com.set("Select")
            templist = []
            for key in self.proccap.keys():
                if key[1]==self.cv_add_src.get():
                    templist.append(key[2])
            templist = list(set(templist))
            templist.sort()
            for k in templist:
                menu.add_command(label=k, command=lambda k=k: self.cv_add_com.set(k))
            if len(templist)==0:
                self.cv_add_com.set("NOT FOUND")
        else:
            self.cv_add_com.set("Select")
            for k in self.commodities:
                menu.add_command(label=k, command=lambda k=k: self.cv_add_com.set(k))

    def load_basket(self):
        '''
        This function loads the default food basket for the current activity
        '''

        i = 0
        for k in self.commodities:
            v = self.baskets[self.ben.get(),k]
            if v > 0:
                self.fix_com[i].set(k)
                self.fix_quant[i].set(v)
                i+=1
        j = len(self.fix_com.keys())
        while i < j:
            self.fix_com[i].set("Select")
            self.fix_quant[i].set("N/A")
            i+=1
        self.set_fb()

    def convert_unit(self, v1, u1, u2):
        '''
        Convert one unit to another (used in update_conversion(self, *args))
        '''

        d, s, c = self.fix_days.get(), self.fix_hh.get(), self.fix_g2l.get()
        if s == 0 or d == 0 or c == 0 :
            return 0  # to prevent zero division
        # first convert to g/p/d
        if u1 == "g/p/d":
            v = v1
        elif u1 == "KG/hh/m":
            v = v1 * 1000 / s / d
        else: # u1 == "L/hh/m"
            v = v1 * c / s / d
        # then convert to u2
        if u2 == "g/p/d":
            v2 = v
        elif u2 == "KG/hh/m":
            v2 = v / 1000 * s * d
        else: # u2 == "L/hh/m"
            v2 = v / c * s * d
        return "{0:.2f}".format(v2)

    def OnFrameConfigure(self,event): # Auxiliary function for the scrollbar
        '''
        Reset the scroll region to encompass the inner frame
        '''

        self.bmcanvas.configure(scrollregion=self.bmcanvas.bbox("all"))

    def getbaseline(self, *args): # Grabs the filter baseline
        '''
        This is a roundabout way of obtaining a reference to Python variables
        '''

        try: # args[0] = PY_VARi , where i is some number (args[0] is the i'th created variable I think)
            self.filter = int(args[0][-2:])
        except: # i had only 1 digit
            self.filter = int(args[0][-1:])

    def fmt_wcommas(self,amount):
        '''
        This function returns the input number with a $#,###.## format
        '''

        if not amount: return '$0.00' # If I got zero return zero (0.00)

        if amount < 0: sign="-" # Handle negative numbers
        else: sign=""

        whole_part=abs(long(amount)) # Split into fractional and whole parts
        fractional_part=abs(amount)-whole_part
        add = ("%.2f" % fractional_part)[0]
        if  add > 0:
            whole_part+=long(add)  # The original function was unable to round up to 1.00

        temp="%i" % whole_part # Convert to string
        digits=list(temp) # Convert the digits to a list

        # Calculate the pad length to fill out a complete thousand and add
        # the pad characters (space(s)) to the beginning of the string.
        padchars=3-(len(digits)%3)
        if padchars != 3: digits=tuple(padchars*[' ']+digits)
        else: digits=tuple(digits)

        sections=len(digits)/3
        mask=sections*",%s%s%s" # Create the mask for formatting the string
        outstring=mask % digits
        outstring=sign+"$"+outstring[1:].lstrip() # Drop off the leading comma and add currency and sign
        outstring+=("%.2f" % fractional_part)[1:] # Add back the fractional part
        return outstring

    def sortTree(self, tree, col, descending):
        '''
        This function sorts the contents of a column when its header is clicked
        '''

        # grab values to sort
        data = [(tree.set(child, col), child) \
            for child in tree.get_children('')]
        # now sort the data in place
        data.sort(reverse=descending, key=self.sortByValue)
        for ix, item in enumerate(data):
            tree.move(item[1], '', ix)
        # switch the heading so it will sort in the opposite direction
        tree.heading(col, command=lambda col=col: self.sortTree(tree, col, \
            int(not descending)))

    def sortByValue(self, TUPLE):
        '''
        This function accepts a tuple and returns a value that can be used for sorting purposes
        '''

        TUPLE = [TUPLE[0],TUPLE[1]] # tuples can't be edited, so we convert it to an editable list
        TUPLE[0] = TUPLE[0].replace(',','') # remove commas
        if TUPLE[0][0] == "$":
            return float(TUPLE[0][1:])
        elif TUPLE[0][-1] == "%":
            return float(TUPLE[0][:-1])
        else:
            for c in TUPLE[0]:
                if c in "0000123456789000-.":
                    numeric = True
                else:
                    numeric = False
                    break
            if numeric == True:
                return float(TUPLE[0])
            else:
                return TUPLE[0]
        # NB: If the column to be sorted contains a string, we return the string; otherwise we return a float (so we need to exclude commas, $, %, etc.)

    def xldate2month(self, DATE, TYPE):
        '''
        This function converts an Excel date to a month name (MMM)
        '''

        d = datetime.datetime(1899, 12, 30) + datetime.timedelta(days=DATE)
        s = d.isoformat() # "YYYY-MM-DD" + "T00:00:00"
        i = s[5:7]
        if i == "01":
            m = "Jan"
        elif i == "02":
            m = "Feb"
        elif i == "03":
            m = "Mar"
        elif i == "04":
            m = "Apr"
        elif i == "05":
            m = "May"
        elif i == "06":
            m = "Jun"
        elif i == "07":
            m = "Jul"
        elif i == "08":
            m = "Aug"
        elif i == "09":
            m = "Sep"
        elif i == "10":
            m = "Oct"
        elif i == "11":
            m = "Nov"
        elif i == "12":
            m = "Dec"
        else:
            m = "Something went wrong :("
        i = s[2:4]
        if TYPE==0:
            return m
        else:
            return m + "-" + i

class ToolTip(object):

    def __init__(self, widget):
        self.widget = widget
        self.tipwindow = None
        self.id = None
        self.x = self.y = 0

    def showtip(self, text):
        "Display text in tooltip window"
        self.text = text
        if self.tipwindow or not self.text:
            return
        x, y, cx, cy = self.widget.bbox("insert")
        x = x + self.widget.winfo_rootx() + 27
        y = y + cy + self.widget.winfo_rooty() +27
        self.tipwindow = tw = Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry("+%d+%d" % (x, y))
        try:
            # For Mac OS
            tw.tk.call("::tk::unsupported::MacWindowStyle",
                       "style", tw._w,
                       "help", "noActivates")
        except TclError:
            pass
        label = ttk.Label(tw, text=self.text, justify=LEFT,
                      background="white", relief=SOLID, borderwidth=1,
                      font=("tahoma", "8", "normal"), foreground="black")
        label.pack(ipadx=1)

    def hidetip(self):
        tw = self.tipwindow
        self.tipwindow = None
        if tw:
            tw.destroy()

def createToolTip(widget, text):
    toolTip = ToolTip(widget)
    def enter(event):
        toolTip.showtip(text)
    def leave(event):
        toolTip.hidetip()
    widget.bind('<Enter>', enter)
    widget.bind('<Leave>', leave)

class TextRedirector(object):
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        if "Error" in str or "error" in str or "Warning" in str or "warning" in str:
            tag = "stderr"
        else:
            tag = self.tag
        self.widget.insert("end", str, (tag,))
        #self.widget.insert("end",self.tag)
        self.widget.yview(END)
        self.widget.configure(state="disabled")
        self.widget.update_idletasks()

class TextRedirector2(object):
    def __init__(self):
        None
    def write(self,str):
        None

class McListBox(object):
    """use a ttk.TreeView as a multicolumn ListBox"""
    def __init__(self,header,data):
        self.tree = None
        self._setup_widgets(header,data)
        self._build_tree(header,data)
    def _setup_widgets(self,header,data):
        s = """\
click on header to sort by that column
to change width of column drag boundary
        """
        msg = ttk.Label(wraplength="4i", justify="left", anchor="n",
            padding=(10, 2, 10, 6), text=s)
        msg.pack(fill='x')
        container = ttk.Frame()
        container.pack(fill='both', expand=True)
        # create a treeview with dual scrollbars
        self.tree = ttk.Treeview(columns=header, show="headings")
        vsb = ttk.Scrollbar(orient="vertical",
            command=self.tree.yview)
        hsb = ttk.Scrollbar(orient="horizontal",
            command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set,
            xscrollcommand=hsb.set)
        self.tree.grid(column=0, row=0, sticky='nsew', in_=container)
        vsb.grid(column=1, row=0, sticky='ns', in_=container)
        hsb.grid(column=0, row=1, sticky='ew', in_=container)
        container.grid_columnconfigure(0, weight=1)
        container.grid_rowconfigure(0, weight=1)
    def _build_tree(self,header,data):
        for col in header:
            self.tree.heading(col, text=col.title(),
                command=lambda c=col: sortby(self.tree, c, 0))
            # adjust the column's width to the header string
##            self.tree.column(col,
##                width = tkFont.Font().measure(col.title()))
        for item in data:
            self.tree.insert('', 'end', values=item)
            # adjust column's width if necessary to fit each value
##            for ix, val in enumerate(item):
##                col_w = tkFont.Font().measure(val)
##                if self.tree.column(header[ix],width=None)<col_w:
##                    self.tree.column(header[ix], width=col_w)
def sortby(tree, col, descending):
    """sort tree contents when a column header is clicked on"""
    # grab values to sort
    data = [(tree.set(child, col), child) \
        for child in tree.get_children('')]
    # if the data to be sorted is numeric change to float
    #data =  change_numeric(data)
    # now sort the data in place
    data.sort(reverse=descending)
    for ix, item in enumerate(data):
        tree.move(item[1], '', ix)
    # switch the heading so it will sort in the opposite direction
    tree.heading(col, command=lambda col=col: sortby(tree, col, \
        int(not descending)))



########################################################################################
################## Executed code #######################################################
########################################################################################

rootWin = Tk()
rootWin.title("AID-M: WFP's Assistant for Integrated Decision-Making")
script_dir = os.path.dirname(os.path.abspath(__file__))
dest_dir = os.path.join(script_dir, 'data')
dest_dir = os.path.join(dest_dir, 'wfp.ico')
rootWin.wm_iconbitmap(dest_dir)
bgcolor = '#%02x%02x%02x' % (51, 128, 255)
rootWin.configure(background=bgcolor)

app = UNWFPModel(rootWin)
rootWin.mainloop()




