import core.tdm as tdm
import core.tdm.trace
import core.eam as eam
from core.tdm.trace import TraceResults, TraceItemEntities
import core.gui as gui
from core.gdm.lookuptables import *
from core.gui import *
from core.gui.editpanel import *
import datetime
import os
import sys
import core.jms
import win32com.client
from win32com.client import constants

# Two lines below generated from the command:
# python makepy.py -i VISLIB.DLL
from win32com.client import gencache
gencache.EnsureModule('{00021A98-0000-0000-C000-000000000046}', 0, 4, 11)
#Test

#===========================================================================================================================
# Return true if the passed in entity is an ISP equipment, otherwise return false
#===========================================================================================================================
def is_isp_class(ent):
	isp_classes = ["ISP_RACK","ISP_PORT_AND_OWNER_mixin","ISP_CABLE", "TERM_PORTGR","FIBER_CABLE_SEG_ISP","COUPLER_PORTGR"]
	if ent is None:
		return False
	for isp_class in isp_classes:
		if ent.is_class(isp_class):
			return True
	return False

#===========================================================================================================================
# Return False if we reach a "stop" class entity or if the entity is null, otherwise return true
#===========================================================================================================================
def is_stop_class(ent):
	stop_classes = ["SPLICE_ENCLOSURE","RF_NODE","fdm_storage_loop"] #Removed SITE from list
	if ent is None:
		return False
	for stop_class in stop_classes:
		if ent.is_class(stop_class):
			return False
	return True
	
#===========================================================================================================================
# Helper function for retreving attributes
#===========================================================================================================================
def checkValue(value):
	if value is None:
		return ""
	else:
		return str(value)


def getChassis(ent):

	parent = ent
	while True:
		try:
			if parent.is_class("ISP_PORT"):
				parent = parent.ISPA_PORT_OWNER_FK
			else:
				parent = parent.PARENT_NODEHOUSING
		except:
			break

		if parent.is_class("ISP_CHASSIS"):
			break

	return parent
	
def addedInJob(ent,label):

	touched = False
	if  ent in core.jms.CurrentJobInfo().AddedInJob(ent.classname()):
		touched =  True
	if touched:
		return label+"_new"
	else:
		return label
	

def main():


	#Set up required dictionaries
	#circuit_state_info = ConfigurationDictionary("FIBER_RING_DICT")
	port_info = ConfigurationDictionary("PORT_DICT")
	equip_dict = ConfigurationDictionary("EQDICT")
	trace_Reports = []
	trace_Reports_Desc = []
	master_circuits = []

	trace_Reports_Desc.append("Customer Site")
	trace_Reports_Desc.append("Customer Site name")
	trace_Reports_Desc.append("Customer Site CLLI")
	trace_Reports_Desc.append("Customer Site type")
	trace_Reports_Desc.append("Customer Site location")
	trace_Reports_Desc.append("Customer Site address")
	trace_Reports_Desc.append("Customer Site End Equip")
	trace_Reports_Desc.append("Customer Site Equip")
	trace_Reports_Desc.append("Customer Site OSP Fiber Cable")
	trace_Reports_Desc.append("OSP De(Mux)")
	trace_Reports_Desc.append("Headend")
	trace_Reports_Desc.append("Headend name")
	trace_Reports_Desc.append("Headend CLLI")
	trace_Reports_Desc.append("Headend type")
	trace_Reports_Desc.append("Headend location")
	trace_Reports_Desc.append("Headend address")
	trace_Reports_Desc.append("Headend End Equip")
	trace_Reports_Desc.append("Headend Equip")
	trace_Reports_Desc.append("Headend OSP Fiber Cable")
	trace_Reports_Desc.append("Master Circuit Name")
	trace_Reports_Desc.append("Job Name")
	trace_Reports_Desc.append("Job Owner")
	trace_Reports_Desc.append("Date")

	
	for result in TraceResults().getTraceResults(): #for each result in the trace report

		entity_list = []
		#===========================================================================================================================
		# Trace helper function
		#===========================================================================================================================
		def storeTraceResult(node,direction,parent):
			entity_list.append(node) #append the nodes to the end of the entity list
		#result.trace_tree.applyBidirectional(core.tdm.trace.TraceNode.printCallback, walk_type = "bidirectional")
		#print "----------"
		result.trace_tree.applyBidirectional(storeTraceResult) #traverse the trace tree in both directions

		if len(entity_list)>1: #if the length of the entity list is more than 1 then

			trace_Report = attributes = [""]*23 #declare trace_reports with 23 empty attributes
			osp_indx_1 = -1                   #osp index 1 set to -1 (for flagging when the circuit has hit OSP fiber)
			osp_indx_2 = len(entity_list)     #osp index 2 set to length of entity list (same purpose as above)
			
			#master circuit details
			master_circuit = None
			if result.segment.is_class("_tdm_hascircuitproperties"):  #checking if the segment of the current result is part of _tdm_hascircuitproperties
				circuit_state = SPATIALnet.service("ndm$property_get_circuitstate", result.segment, result.sequence)  #retrieve the circuit state from spatialnet service
					
				if result.channel: #if the result is channelized
					for ch in circuit_state.fdm_sub_channel_scan: #loop through all the channels and if the the channel matches the circuit state channel then the circuit ID is on that channel
						if ch.fdm_ring_sequence == result.channel:
							master_circuit = ch.fdm_ringmaster_fk
							break
						
				if master_circuit is None:   #if there is no master circuit then
					master_circuit = circuit_state.fdm_ringmaster_fk   #set the master circuit to the circuit ID in the circuit state
				if master_circuit:   #if there is a master circuit then
						mc = str(master_circuit.fdm_ringmaster_name) # set mc = to the ring name (CLFI) after making it a string
						trace_Report[19] = mc #put the master circuit ID into attribute 19 in the trace report
						if master_circuits.count(mc)==0 and mc is not None and len(mc)>0: #if there are no master circuits and mc is not empty and the length of mc is greater than 0 then
							master_circuits.append(mc)  #append mc to the end of master_circuits
			else: #if the segment of the current result is not part of _tdm_hascircuitproperties then
				if len(master_circuits)>0: #check if the length of master_circuits is greater than 0
					trace_Report[19] = master_circuits[len(master_circuits)-1]  #make the last master circuit added the attribute 

			#JOB NAME
			trace_Report[20] =  str(eam.current_job().jms_job_description)

			#JOB NAME
			trace_Report[21] =  str(eam.current_job().eam_job_owning_user.scm_real_name)
			
			#current date
			trace_Report[22] = str(datetime.datetime.now().date())


			a_end_isp_design = ""
			a_end_equip = []
			a_end_osp_cable = ""
			osp_equip = None
			correct_order = True
			first_port = None
			
			customer_site_indx = 0
			headend_site_indx = 10

			for i in range(len(entity_list)): #loop through the entity list
			
				ent2 = entity_list[i].entity #take the ith entity and put it into ent2
				
				if ent2.is_class("ISP_PORT"): #if the current entity is an ISP_PORT
				
					#print str(entity_list[i].entity)+ " : depth("+ str(entity_list[i].depth)+"), branch("+str(entity_list[i].branch_number)+")"
				
					if first_port is None: #and if the first port has not yet been found then
						first_port = ent2  #make the ith entity the first port

					parent = ent2.ISPA_PORT_OWNER_FK #find the parent of the ent2

					if parent.fdm_interface_fk is not None: #if the parent has an osp interface then
						#found patch panel
						chassis = getChassis(ent2)  #get the chassis of ent2
						pnl = checkValue(chassis.ISPA_NAME) + " ; "+checkValue(chassis.ISPA_SECTION_F_CODE) + " ; " + checkValue(ent2.ISPA_PORT_NAME) 
						a_end_equip.append(addedInJob(chassis,"Patch Panel")+": "+pnl) #get the name of the panel and all information

					#check for De(Mux) or patch panel
					else:     #if the parent has no osp interface
						#Dictionary look up for isp equipment
						try:
							equip_type_details = equip_dict.values(parent.ISPA_EQUIP_DICT_FK.NETWORK_KEY)  #get the details of the type of equipment that is ent2
							desc = checkValue(equip_type_details.DESC1) #get the description of the chassis from the dictionary

							if desc.upper().find("MULTIPLEXER")!= -1: #if there is the substring MULTIPLEXER in the description of the chassis then
								#found mux! Now get the information of the mux shelf
								mux = checkValue(equip_type_details.MODEL) + " ; " + desc + " ; "+ checkValue(parent.ISPA_SECTION_F_CODE) + " ; "+ checkValue(parent.ISPA_NAME)

								if a_end_equip.count(addedInJob(parent,"Mux")+": "+mux)==0: #if the parent and the mux were both added in this job then
									a_end_equip.append(addedInJob(parent,"Mux")+": "+mux)      #add the parent chassis to the a_end_equipment
							else:
								#found true end
								if ent2 != first_port and entity_list[i].branch_number==1: #if ent2 is not the first port and the ith element in the entity list's branch # = 1 then
									correct_order=False #the entities are not in the correct order
									         #add all information for the isp a end
								a_end_isp_design = addedInJob(ent2,"End Equipment")+": "+checkValue(ent2.ISPA_SECTION_F_CODE) + " ; " + checkValue(parent.ISPA_NAME) + " ; "+equip_type_details.MODEL + " ; " + equip_type_details.DESC1

						except Exception as e:
							#lov conversion not found
							print e

				elif ent2.is_class("COUPLER_PORTGR"): #if the a end is not a mux then check if it a demux so check if it is from the class COUPLER_PORTGR
					#check for De(Mux)
					coupler = ent2.ndm_port_owner  #if ent2 is in the coupler class then put its parent in coupler
					isp_rack = coupler.PARENT_NODEHOUSING  #put the housing of the coupler (rack) into isp_rack

					if len(a_end_osp_cable) ==0: #if the length of the list of osp cables is 0 (no osp cables have been discovered yet) then
						if ent2.is_class("_tdm_hascircuitproperties"): #check if ent2 has circuit properties
							circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",entity_list[i+1].entity,entity_list[i+1].sequence)#get the circuit state of the next entity in the list 
							a_end_osp_cable = checkValue(circuit_state.fdm_usage_desc) #get the description of the circuit state
					
					if isp_rack.is_class("ISP_RACK"):  #if isp_rack is a rack then
						is_wdm = coupler.fdm_equip_type_code.upper().find("WDM_10WAY") != -1 #flag if coupler has WDM_10WAY in its description 
						if (is_wdm and entity_list[i].sequence!= 1) or (not is_wdm and entity_list[i].sequence==3): #if the coupler is wdm the ith entity has a sequence not equal to 1 OR coupler is not wdm and ith entity has sequence = 3 then put all information in mux  
							mux = checkValue(coupler.fdm_equip_name) + " ; " + checkValue(coupler.fdm_equip_type_code) + " ; " + checkValue(isp_rack.ISPA_NAME) + " ; " + checkValue(isp_rack.ISPA_SECTION_F_CODE)
				
							if ent2.is_class("_tdm_hascircuitproperties"): #if ent2 has circuit properties then
								circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",ent2,entity_list[i].sequence) #get its circuit state
								a_end_equip.append(addedInJob(ent2,"Multi-Fiber Cable")+": "+checkValue(circuit_state.fdm_usage_desc)) #append the multifiber cable added in the job

							a_end_equip.append(addedInJob(coupler,"Mux")+": "+mux) #append the mux added in the job
								
							if ent2.is_class("_tdm_hascircuitproperties"):
								circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",entity_list[i+1].entity,entity_list[i+1].sequence)
								a_end_equip.append(addedInJob(entity_list[i+1].entity,"Multi-Fiber Cable")+": "+checkValue(circuit_state.fdm_usage_desc))


				#check for patch cord
				elif ent2.is_class("ISP_PATCH_CORD"): #if the current entity is a patch cord then

					desc = "Length: "+ str(ent2.LE_LENGTH) + " ; " #get the length of the patch cord
					
					#Dictionary look up for isp equipment
					try:
						equip_type_details = equip_dict.values(ent2.ISPA_EQUIP_DICT_FK.NETWORK_KEY)
						desc = desc + equip_type_details.MODEL
					except Exception as e:
						#lov conversion not found
						#pass
						print (str(e))
					
					
					a_end_equip.append(addedInJob(ent2,"Patch Cable")+": "+desc) #append the patch cable to a_end_equipment
						
					
				#if found FIBER_CABLE_SEG_UNCON, check connections on both sides to ensure it leaves the building.
				#if not, this could be a data issue
				
				elif ent2.is_class("FIBER_CABLE_SEG_UNCON"): #check if the current entity is a fiber segment
				
					try:
						owner = ent2.ndm_leseg_owner #get owner
						start_joint_parent = owner.ndm_le_startjoint.PARENT_NODEHOUSING #get start joint
						end_joint_parent = owner.ndm_le_endjoint.PARENT_NODEHOUSING  #get end joint
						
						if not (is_isp_class(start_joint_parent) and is_isp_class(end_joint_parent)): #if both joints are not in the isp then
							osp_indx_1 = i #make i the osp index 1 and break out of loop (entity i is where osp starts)
							break
					except Exception as e:
						osp_indx_1 = i
						break

				elif not (is_isp_class(ent2)): #if the ith entity is not in the isp layer then its in osp and break.
					osp_indx_1 = i
					break
					
				osp_indx_1 = i
				
			#flip the equip list if trace is not in correct order
			if not correct_order:
				a_end_equip.reverse()
				
			a_end_nh = entity_list[0].upstream_osp_nh			
			a_end_key = a_end_nh.NETWORK_KEY
			a_end_address = "%s ; %s ; %s ; %s" % (a_end_nh.fdm_address1, a_end_nh.fdm_town,a_end_nh.fdm_state,a_end_nh.fdm_zipcode)
			a_end_name = a_end_nh.fdm_designation
			a_end_clli = a_end_nh.gdm_ea_attr_01
			a_end_type = a_end_nh.fdm_site_type_code
			a_end_location = a_end_nh.fdm_nh_location
			# print ("A/Z ends points for Master Circuit->",master_circuit)

			if a_end_type.lower() == "headend":
				customer_site_indx = 10
				headend_site_indx = 0
			elif a_end_type.lower() == "customer":
				customer_site_indx = 0
				headend_site_indx = 10
		

			z_end_isp_design = ""
			z_end_equip = []
			z_end_osp_cable = ""
			correct_order = True
			first_port = None

			for r in reversed(xrange(len(entity_list))): #start from the last entity in the list and loop

				ent2 = entity_list[r].entity #make ent2 the rth entity in the entity list
						
				if ent2.is_class("ISP_PORT"): #if ent2 is an isp port then
				
					#print str(entity_list[r].entity)+ "(z) : depth("+ str(entity_list[r].depth)+"), branch("+str(entity_list[r].branch_number)+")"
				
					if first_port is None: #if the first port has not been discovered yet then ent2 is the first port
						first_port = ent2

					parent = ent2.ISPA_PORT_OWNER_FK #get ent2's parent
					#check for patch panel
					
					if parent.fdm_interface_fk is not None: #if there is some osp interface then
						chassis = getChassis(ent2)  #get the chassis of ent2 (it is a panel)
						pnl = checkValue(chassis.ISPA_NAME) + " ; "+checkValue(chassis.ISPA_SECTION_F_CODE) + " ; " + checkValue(ent2.ISPA_PORT_NAME)  #put all information into pnl
						z_end_equip.append(addedInJob(chassis,"Patch Panel")+": "+pnl) #append pnl into z end equipment

					#check for De(Mux) or patch panel
					else:
						#Dictionary look up for isp equipment
						try:
							equip_type_details = equip_dict.values(parent.ISPA_EQUIP_DICT_FK.NETWORK_KEY)
							desc = checkValue(equip_type_details.DESC1)
							if desc.upper().find("MULTIPLEXER")!= -1:
								#found mux
								mux = checkValue(equip_type_details.MODEL) + " ; " + desc + " ; "+ checkValue(parent.ISPA_SECTION_F_CODE) + " ; "+ checkValue(parent.ISPA_NAME)
								if z_end_equip.count(addedInJob(parent,"Mux")+": "+mux)==0:
									z_end_equip.append(addedInJob(parent,"Mux")+": "+mux)
							else:
								#found true end
								if ent2 != first_port and entity_list[r].branch_number==1:
									correct_order=False
									
								z_end_isp_design = addedInJob(ent2,"End Equipment")+": "+checkValue(ent2.ISPA_SECTION_F_CODE) + " ; " + checkValue(parent.ISPA_NAME) + " ; "+equip_type_details.MODEL + " ; " + equip_type_details.DESC1

						except Exception as e:
							#lov conversion not found
							print "z: "+str(e)

				elif ent2.is_class("COUPLER_PORTGR"):
					#check for De(Mux)
					coupler = ent2.ndm_port_owner
					isp_rack = coupler.PARENT_NODEHOUSING

					if len(z_end_osp_cable) ==0:
						if ent2.is_class("_tdm_hascircuitproperties"):
							circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",entity_list[r-1].entity,entity_list[r-1].sequence)
							a_end_osp_cable = checkValue(circuit_state.fdm_usage_desc)
					
					if isp_rack.is_class("ISP_RACK"):
						is_wdm = coupler.fdm_equip_type_code.upper().find("WDM_10WAY") != -1
						if (is_wdm and entity_list[r].sequence!= 1)  or (not is_wdm and entity_list[r].sequence==3):
							mux = checkValue(coupler.fdm_equip_name) + " ; " + checkValue(coupler.fdm_equip_type_code) + " ; " + checkValue(isp_rack.ISPA_NAME) + " ; " + checkValue(isp_rack.ISPA_SECTION_F_CODE)

							if ent2.is_class("_tdm_hascircuitproperties"):
								circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",ent2,entity_list[r].sequence)
								z_end_equip.append(addedInJob(ent2,"Multi-Fiber Cable")+": "+checkValue(circuit_state.fdm_usage_desc))

							z_end_equip.append(addedInJob(coupler,"Mux")+": "+mux)
								
							if ent2.is_class("_tdm_hascircuitproperties"):
								circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",entity_list[r-1].entity,entity_list[r-1].sequence)
								z_end_equip.append(addedInJob(entity_list[r-1].entity,"Multi-Fiber Cable")+": "+checkValue(circuit_state.fdm_usage_desc))


				#check for patch cord
				elif ent2.is_class("ISP_PATCH_CORD"):

					desc = "Length: "+ str(ent2.LE_LENGTH) + " ; "
					
					#Dictionary look up for isp equipment
					try:
						equip_type_details = equip_dict.values(ent2.ISPA_EQUIP_DICT_FK.NETWORK_KEY)
						desc = desc + equip_type_details.MODEL
					except Exception as e:
						#lov conversion not found
						#pass
						print (str(e))

					
					z_end_equip.append(addedInJob(ent2,"Patch Cable")+": "+desc)
					
					
				#if found FIBER_CABLE_SEG_UNCON, check connections on both sides to ensure it leaves the building.
				#if not, this could be a data issue
				
				elif ent2.is_class("FIBER_CABLE_SEG_UNCON"):
				
					try:
						owner = ent2.ndm_leseg_owner
						start_joint_parent = owner.ndm_le_startjoint.PARENT_NODEHOUSING
						end_joint_parent = owner.ndm_le_endjoint.PARENT_NODEHOUSING
						
						if not (is_isp_class(start_joint_parent) and is_isp_class(end_joint_parent)):
							osp_indx_2 = r
							break
						
					except Exception as e:
						osp_indx_2 = r
						break	

				elif not (is_isp_class(ent2)):
					osp_indx_2 = r
					break

				osp_indx_2 = r
				
			#flip the equip list if trace is not in correct order
			if not correct_order:
				z_end_equip.reverse()
					
			z_end_nh = entity_list[len(entity_list)-1].upstream_osp_nh
			z_end_key = z_end_nh.NETWORK_KEY
			z_end_address = "%s ; %s ; %s ; %s" % (z_end_nh.fdm_address1, z_end_nh.fdm_town,z_end_nh.fdm_state,z_end_nh.fdm_zipcode)
			z_end_name = z_end_nh.fdm_designation
			z_end_clli = z_end_nh.gdm_ea_attr_01
			z_end_type = z_end_nh.fdm_site_type_code
			z_end_location = z_end_nh.fdm_nh_location
			
			#check a/z end addresses
			add_a_end = True   #if both of these are true then add customer and headend locations
			add_z_end = True
			
			#if z_end_address == a_end_address: #if the trace is within the same building and
			#	if z_end_type.upper() == "CUSTOMER": # if z end is the customer end
			#		add_a_end=False    #dont add the a side
			#	else:        #otherwise
			#		add_z_end=False   #don't add z end
					
					
			#check for z_end_equip ending with multi-fiber cable
            #checking both ends if the last thing appended was a multifiber cable
			if len(a_end_equip)>0:
				if a_end_equip[len(a_end_equip)-1].find("Multi-Fiber") != -1:
					a_end_equip.pop(len(a_end_equip)-1)
			
			if len(z_end_equip)>0:
				if z_end_equip[len(z_end_equip)-1].find("Multi-Fiber") != -1:
					z_end_equip.pop(len(z_end_equip)-1)

			if add_z_end: #if the add a end flag is true then add 
				trace_Report[headend_site_indx] = str(z_end_nh)
				trace_Report[headend_site_indx+1] = str(z_end_name)
				trace_Report[headend_site_indx+2] = str(z_end_clli)
				trace_Report[headend_site_indx+3] = str(z_end_type)
				trace_Report[headend_site_indx+4] = str(z_end_location)
				trace_Report[headend_site_indx+5] = str(z_end_address)
				trace_Report[headend_site_indx+6] = str(z_end_isp_design)
				trace_Report[headend_site_indx+7] = z_end_equip #z_end_osp_design
			
			if add_a_end:
				trace_Report[customer_site_indx] = str(a_end_nh)
				trace_Report[customer_site_indx+1] = str(a_end_name)
				trace_Report[customer_site_indx+2] = str(a_end_clli)
				trace_Report[customer_site_indx+3] = str(a_end_type)
				trace_Report[customer_site_indx+4] = str(a_end_location)
				trace_Report[customer_site_indx+5] = str(a_end_address)
				trace_Report[customer_site_indx+6] = str(a_end_isp_design)
				trace_Report[customer_site_indx+7] = a_end_equip #a_end_osp_design
						
			osp_mux = ""
			#iterate through OSP equipment

			if len(a_end_osp_cable)==0:  #if a_end_osp_cable is empty
				for a in range(osp_indx_1+1,osp_indx_2): #loop from osp index 1 to osp index 2
					ent2 = entity_list[a].entity  #get the ath entity from the entity list
					if ent2.is_class("FIBER_CABLE_SEG_UNCON"): #if that entity is a fiber cable segment then
						if ent2.is_class("_tdm_hascircuitproperties"): # if that entity has a circuit ID
							circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",ent2,entity_list[a].sequence) #get the circuit state for that entity
							a_end_osp_cable = checkValue(circuit_state.fdm_usage_desc) # make that the a_end_osp_cable
							break #get out of for loop

			if len(z_end_osp_cable)==0: #if the z_end_osp_cable is empty
				for a in reversed(xrange(osp_indx_1+1,osp_indx_2+1)):#loop from osp index 2 (add 1) down to osp index 1 (add 1)
					ent2 = entity_list[a].entity #get the ath entity from the entity list
					if ent2.is_class("FIBER_CABLE_SEG_UNCON"): # if the entity is an osp fiber cable segment then 
						if ent2.is_class("_tdm_hascircuitproperties"): # if the entity has a circuit ID then 
							circuit_state = SPATIALnet.service("ndm$property_get_circuitstate",ent2,entity_list[a].sequence) #get the circuit state for that entity
							z_end_osp_cable = checkValue(circuit_state.fdm_usage_desc) # make that the z end osp cable
							break #get out of for loop

			#look for any De(mux) located OSP

			for a in range(osp_indx_1+1,osp_indx_2): #loop from osp index 1 (add 1) to osp index 2
				ent2 = entity_list[a].entity  #get the ath entity

				if ent2.is_class("COUPLER_PORTGR"): #if the entity is a coupler port
					coupler = ent2.ndm_port_owner #set coupler to the owner of that port
					splice_case = coupler.PARENT_NODEHOUSING # set splice_case to parent of the coupler

					is_10waywdm = coupler.fdm_equip_type_code.upper().find("WDM_10WAY") != -1 #create condidtion (if the coupler is a 10 way wdm)
					is_wdm = coupler.fdm_equip_type_code.upper().find("WDM") != -1  #create condidtion (if the coupler is a wdm)
					
					if is_10waywdm or (not is_10waywdm and is_wdm and entity_list[a].sequence==3): #if the coupler is a wdm10way OR is not a wdm10way but is a wdm and the entity has sequence = 3
						if splice_case.is_class("SPLICE_CASE"): # if splice_case is a splice case then
							osp_mux = addedInJob(ent2,"OSP Mux")+": "+checkValue(coupler.fdm_equip_name) + " ; " + checkValue(coupler.fdm_equip_type_code) + " ; " + checkValue(splice_case.fdm_nh_location) + " ; " + checkValue(splice_case.fdm_designation) + " ; " + \
										checkValue(splice_case.fdm_address1) + " " + checkValue(splice_case.fdm_town) + " " + checkValue(splice_case.fdm_zipcode) #get the information from the mux, coupler and splice case and put it into osp_mux
							break

			trace_Report[customer_site_indx+8] = str(a_end_osp_cable) #whcihever index is customer, add 8 to get to the field for osp cable and put the string of a_end_osp_cable
			trace_Report[headend_site_indx+8] = str(z_end_osp_cable)  #whcihever index is headend, add 8 to get to the field for osp cable and put the string of z_end_osp_cable
			trace_Report[9] = str(osp_mux) #add the string for osp mux into the 9th attribute in the trace report

			
			flag = True
			index = 0
			direction = -1
			
			if trace_Report[0] == trace_Report[10] and trace_Report[0].find("ISP_BUILDING")==-1: #if a site and z siet are equal and if a site is an isp building then make the flag false (don't add the trace)
				flag = False
			else: #if the trace is not within the same building then
				for r in trace_Reports:  #for each report in trace report
					cond1 = r[6] == trace_Report[6] and r[16] == trace_Report[16]#end equip
					cond2 = r[6] == trace_Report[16] and r[16] == trace_Report[6]#end quip
					cond3 = r[0] == trace_Report[0] and r[10] == trace_Report[10]#site
					cond4 = r[0] == trace_Report[10] and r[10] == trace_Report[0]#site
					cond5 = r[19] == trace_Report[19]
					
					if cond5: #if the master circuit is the same between the rth trace report and the current one
						if cond3 or cond4: #if the sites are the same  between the rth trace report and the current one
							if cond1 or cond2: #if the equipment has the same name then don't add the trace
								flag=False
								break
							elif len(trace_Report[6])>len(r[6]) or len(trace_Report[16])>len(r[16]): #if the name of the equipment in the current trace is longer on either end then set direction and don't append the trace 
								direction = 1
								flag=False
								break
							elif len(trace_Report[6])<=len(r[6]) or len(trace_Report[16])<=len(r[16]):#if the name of the equipment in the current trace is shorter or the same on either side then don't append  
								flag=False
								break

					index = index+1 #increase index

			if direction ==1: #if direction is set then
				trace_Reports.pop(index) #take out trace at position index
				trace_Reports.insert(index,trace_Report) #add current trace in position index
			
			if flag: #if flag is not false then append the trace report
				trace_Reports.append(trace_Report)

	#print len(trace_Reports)
	#Merge Rx & Tx
	report_merged = []
	if len(master_circuits)>1 and len(trace_Reports)>1: # if the length of both master circuits and trace reports are greater than 1 then
		report_merged = trace_Reports[0]  #set report merged to first trace report
		
		for tr_indx in range(1,len(trace_Reports)): #loop through all trace reports
			counter=0
			trace_Report = trace_Reports[tr_indx] # set trace report to the ith trace report
			for r in trace_Report: #for each attribute in the trace report
				if counter==7 or counter==17: # if iteration is on a end or z end equipment then
										
					if len(report_merged[counter])==0: #if equipment has not yet been filled in then fill it in with current equipment name
						report_merged[counter] = r
					elif len(r)==0:  #if not and the current report has no equipment name then move on
						pass
					else: # if they are both not empty then
						c = 0
						while len(r)!=len(report_merged[counter]): #do while the length of current report equipment name and merged report equipment name are not equal
							if len(r) < len(report_merged[counter]): #if the length of r is less than the length of the report merged then
								r.insert(c,report_merged[counter][c]) #insert report merged[counter] at index c into r at index c
							elif len(r) > len(report_merged[counter]): #if the length of r is greater than report merged then
								report_merged[counter].insert(c,r[c])  #insert r at index c into report merged at index c 
								
							c = c+1 #increase c
							
						for i in r:
							indx = r.index(i) #get the index of r where i occurs        
							#if len(report_merged[counter])<indx+1:
							#	report_merged[counter].append(r[indx])
								
							if report_merged[counter][indx] != r[indx]: #if the index does not point to the same thing then
								tuple = r[indx].partition(": ") # partition the indexed item into a 3 tuple (before seperator, seperator, after seperator)
								report_merged[counter][indx] = report_merged[counter][indx]+ "\n\n"+ tuple[2] #add the seperator to the end of the indexed item in report_merged[counter]
				else: # if counter is anything else other than 7 or 17 (not equipment names) then
					if report_merged[counter] != r: #if current report merged and r are not equal then
						report_merged[counter] = report_merged[counter] + " \n\n " +r  # add r to report_merged[counter]
				counter = counter+1 #increase counter
					
		trace_Reports = []
		trace_Reports.append(report_merged) #put reports_merged in trace_Reports

	# for report in trace_Reports:
		# counter = 0
		# for r in report:
			# print trace_Reports_Desc[counter]+ ": "+str(r)
			# counter = counter+1


	return trace_Reports


class DynamicSchemaGenerator:
	
	def __init__(self):
		self.page = None
		self.connectorMaster = None
		self.stencilShapeList = None
		self.doc = None
		self.gap = 1
	
		self.left = {"previousShape": None, "firstShape": None, "x": 2, "y": 3.5, "connectionText": None}
		self.right = {"previousShape": None, "firstShape": None, "x": 12, "y": 3.5, "connectionText": None}
	
	def should_overwrite_file(self, filename):
		"""
		Return a value indicating if we should overwrite this file.

		This should prompt the user (if required) to overwrite the report if
		it already exists. It should return whether to override the report; 
		if it is not to be overwritten, processing will stop.

		"""
		message = "The file %s already exists.\n\n" \
				"Do you wish to overwrite this file?" % filename
		should_overwrite = SPATIALnet.service("gui$prompt_to_continue", message)
		return should_overwrite

	def generateVisio(self, schemaData):
	
		try:
			appVisio = win32com.client.Dispatch("Visio.Application")
			appVisio.Visible = 1
		except:
			sys.exit("Please ensure Microsft Visio Professional 2010 (v.14.0) is installed.")
		
		
		try:

			# open template document and get first page
			# 
			# template is using macros
			# tools -> trust center -> disable all macros except digitally signed macros
			#
			filename = eam.editbuffer("tdm_tm_report_script")
			filename = "\\".join(filename.split('\\')[:-1]) + "\\dynamic.vst"
			doc = appVisio.Documents.Add(filename)
			self.doc=doc
			self.stencilShapeList = appVisio.Documents("RBS_Stencil.vss")
			self.connectorMaster = appVisio.Application.ConnectorToolDataObject
			self.page = doc.Pages.Item(1)
		
		
			# -----------------------
			# outside demux if exists

			# add OutsideOspMux, or if it`s missing join 2 sides
			shape = None
			if schemaData[0].OspMux != "" and schemaData[0].OspMux is not None:
				tuple = schemaData[0].OspMux.partition(": ")
					
				shape = self.page.Drop(self.stencilShapeList.Masters(tuple[0]), 8, 3)
				shape.Cells("Prop.Row_1").Formula = '"' + tuple[2] + '"'
			else:
				shape = self.page.Drop(self.stencilShapeList.Masters("Empty"), 8, 3)
			
			self.left["firstShape"] = shape
			self.left["previousShape"] = shape
			
			self.right["firstShape"] = shape
			self.right["previousShape"] = shape
			
			
			# --------------------
			# left column, headend
			
			# put the outside cable, if there
			self.left["connectionText"] = schemaData[0].HeadendOspFiberCable
			
			# put all other site equipment
			for i in reversed(range(0, len(schemaData[0].HeadendEquip))):
				tuple = schemaData[0].HeadendEquip[i].partition(": ")
				
				# if can't get type from string use cable (blank icon)
				if (len(tuple[2]) > 0):
					self._placeItem(self.left, tuple[0], tuple[2])
				else:
					self._placeItem(self.left, "Unknown", tuple[0])
			
			# put the end equipment
			if (schemaData[0].HeadendEndEquip != None and schemaData[0].HeadendEndEquip != ""):
				tuple = schemaData[0].HeadendEndEquip.partition(": ")
				if tuple[0]== 'End Equipment':
					self._placeItem(self.left, "Router", tuple[2])
				else:
					self._placeItem(self.left, "Router_new", tuple[2])
			
			
			# ---------------------
			# right column, customer
			
			# put the outside cable, if there
			self.right["connectionText"] = schemaData[0].CustomerSiteOspFiberCable
			
			# put all other site equipment
			for i in reversed(range(0, len(schemaData[0].CustomerSiteEquip))):
				tuple = schemaData[0].CustomerSiteEquip[i].partition(": ")
				
				# if can't get type from string use cable (blank icon)
				if (len(tuple[2]) > 0):
					self._placeItem(self.right, tuple[0], tuple[2])
				else:
					self._placeItem(self.right, "Unknown", tuple[0])
			
			# put the end equipment
			if (schemaData[0].CustomerSiteEndEquip != None and schemaData[0].CustomerSiteEndEquip != ""):
				tuple = schemaData[0].CustomerSiteEndEquip.partition(": ")
				if tuple[0]== 'End Equipment':
					self._placeItem(self.right, "Media Converter", tuple[2])
				else:
					self._placeItem(self.right, "Media Converter_new", tuple[2])
			
			
			#populate form
			oleObjects = self.page.OLEObjects
			for intCounter in range(1, oleObjects.Count+1) :
				oleObject = oleObjects(intCounter).Object
				
				if oleObject.Name == "lbl_title":
					oleObject.Caption = schemaData[0].JobName
				elif oleObject.Name == "lbl_drawn_by":
					oleObject.Caption = schemaData[0].JobOwner
				elif oleObject.Name == "lbl_drawn_by_date":
					oleObject.Caption = schemaData[0].Date
				elif oleObject.Name == "lbl_circuit_id":
					oleObject.Caption = schemaData[0].MasterCircuitName
				elif oleObject.Name == "lbl_cust_addr":
					oleObject.Caption = schemaData[0].CustomerSiteName + " ; " + schemaData[0].CustomerSiteCLLI + " ; " +schemaData[0].CustomerSiteAddress
				elif oleObject.Name == "lbl_head_addr":
					oleObject.Caption = schemaData[0].HeadendName + " ; " + schemaData[0].HeadendCLLI + " ; " +schemaData[0].HeadendAddress
					
					
		except Exception as e:
			print e
		
		finally:
			appVisio.Visible = 1
			#setup output file and required handles
			filename = eam.editbuffer("tdm_tm_outputfile")
			#tdm_tm_report_script
			if filename:
				status = SPATIALnet.service("cmn$access", filename, 0)
				if not status:
					if not self.should_overwrite_file(filename):
						return
				if doc:
					self.doc.SaveAs(filename)
		
			
		
	# end of generateVisio function


	# ---------------
	# private methods
	
	def _drawConnection(self, sideData, toShape):
		fromShape = sideData["previousShape"]
		connector = None
		if fromShape is not None and toShape is not None:
			connector = self.page.Drop(self.connectorMaster, 0, 0)		
			
			if sideData["connectionText"] != None and sideData["connectionText"] != "":
				connector.Text = sideData["connectionText"]
				sideData["connectionText"] = None
				connector.cellsU("LineColor").Formula= sideData["connectionTextColor"]
				connector.cellsU("Char.Color").Formula= sideData["connectionTextColor"]
			
			connector.Cells("BeginX").GlueTo(fromShape.Cells("PinX"))
			connector.Cells("EndX").GlueTo(toShape.Cells("PinX"))
			
		return connector
	# end of _drowConnection function
	
	def _placeItem(self, sideData, type, value):
		shape = None
		
		if type == "Multi-Fiber Cable" or type == "Patch Cable":
			if sideData==self.left:
				self.left["connectionText"] = value
				self.left["connectionTextColor"] = "0"
			else:
				self.right["connectionText"] = value
				self.right["connectionTextColor"] = "0"
				
		elif type == "Multi-Fiber Cable_new" or type == "Patch Cable_new":
			if sideData==self.left:
				self.left["connectionText"] = value
				self.left["connectionTextColor"] = "RGB(0,0,255)"
			else:
				self.right["connectionText"] = value
				self.right["connectionTextColor"] = "RGB(0,0,255)"

		else:
			shape = self._placeEquipment(sideData,type, value)
		
		if sideData["firstShape"] is None and shape is not None:
			sideData["firstShape"] = shape
	# end of _placeItem function
	
	
	
	def _placeEquipment(self, sideData,type, value):
		shape = self.page.Drop(self.stencilShapeList.Masters(type), sideData["x"], sideData["y"])
		shape.Cells("Prop.Row_1").Formula = '"' + value + '"'
		sideData["connectionTextColor"] = "0"
		self._drawConnection(sideData, shape)
		sideData["previousShape"] = shape
		sideData["y"] = sideData["y"] + self.gap

		return shape

	# end of _placePatchPanel function

# end of DynamicSchemaGenerator class

class SchemaData:
	CustomerSite = None
	CustomerSiteName = None
	CustomerSiteCLLI = None
	CustomerSiteType = None
	CustomerSiteLocation = None
	CustomerSiteAddress = None
	CustomerSiteEndEquip = None
	CustomerSiteEquip = None
	CustomerSiteOspFiberCable = None
	OspMux = None
	HeadendSite = None
	HeadendName = None
	HeadendCLLI = None
	HeadendType = None
	HeadendLocation = None
	HeadendAddress = None
	HeadendEndEquip = None
	HeadendEquip = None
	HeadendOspFiberCable = None
	MasterCircuitName = None
	JobName = None
	JobOwner = None
	Date = None
	
	def parseArray(self, dataArray):
		result = []
		for i in range(0, len(dataArray)):
			data = SchemaData()
			
			data.CustomerSite = dataArray[i][0]
			data.CustomerSiteName = dataArray[i][1]
			data.CustomerSiteCLLI = dataArray[i][2]
			data.CustomerSiteType = dataArray[i][3]
			data.CustomerSiteLocation = dataArray[i][4]
			data.CustomerSiteAddress = dataArray[i][5]
			data.CustomerSiteEndEquip = dataArray[i][6]
			data.CustomerSiteEquip = dataArray[i][7]
			data.CustomerSiteOspFiberCable = dataArray[i][8]
			data.OspMux = dataArray[i][9]
			data.HeadendSite = dataArray[i][10]
			data.HeadendName = dataArray[i][11]
			data.HeadendCLLI = dataArray[i][12]
			data.HeadendType = dataArray[i][13]
			data.HeadendLocation = dataArray[i][14]
			data.HeadendAddress = dataArray[i][15]
			data.HeadendEndEquip = dataArray[i][16]
			data.HeadendEquip = dataArray[i][17]
			data.HeadendOspFiberCable = dataArray[i][18]
			data.MasterCircuitName = dataArray[i][19]
			data.JobName = dataArray[i][20]
			data.JobOwner = dataArray[i][21]
			data.Date = dataArray[i][22]
			result.append(data)
		return result
#end of SchemaData class

if __name__ == '__main__':
	dsg = DynamicSchemaGenerator()
	dsg.generateVisio(SchemaData().parseArray(main()))
