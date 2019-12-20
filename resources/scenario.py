from flask import Flask, request
from flask_restful import Resource, Api, abort
# from model import WEAP_Data as WEAP_model
from model import WEAP_Simple_Model as WEAP_model
import win32com.client
from model import weap_backend, leap_backend
from model.database import Scenarios, session
import json
class Scenario(Resource):

	def __init__(self):
		pass

	def get(self, pid, model, sid):
		sid = int(sid)

		# Chi: implement a function that gets a scenario based on the 'model' and the 'sid'
		# The returned dict includes all the input and output variables

		flow, timeRange = WEAP_model.get_WEAP_flow_value()
		value = flow[list(flow.keys())[sid]]
		print(value)
		return {
			       'sid': sid,
			       'name': list(flow.keys())[sid],
			       'runStatus': 'finished',
			       'timeRange': timeRange,
			       'numTimeSteps': timeRange[1] - timeRange[0],
			       'var': {'input': '1230',
			               'output': value}
		       }, 200
	def post(self, pid, model, sid):
		data = request.get_json(self)
		print(data)
		win32com.CoInitialize()
		WEAP = win32com.client.Dispatch('WEAP.WEAPApplication')
		start_year = WEAP.BaseYear
		end_year = WEAP.EndYear
		WEAP.Calculate()
		win32com.CoUninitialize()
		return {"response": "POST method was just called!"}, 201

class ScenarioList(Resource):

	def __init__(self):
		pass

	def get(self, pid, model):
		# Chi: implement a function that gets all the existing scenarios
		# and return their brief summary as follows:
		scenario_list = []
		scenarios = ['Reference', '5% Population Growth', '10% Population Growth']
		for i in range(3):
			# get the object of the scenario here

			year = [1986, 2008]

			value = WEAP_model.get_WEAP_flow_value()
			scenario_summary = {
				'sid': i,
				'name':   scenarios[i],
				'runStatus': 'finished',
				'timeRange': year,
				'numTimeSteps': (year[1] - year[0]),
				'__filled': False
			}

			scenario_list.append(scenario_summary)

		return scenario_list, 200

		# return [{
		#     'sid': 'name of the scenario',
		#     'runStatus': 'finished',
		#     'timeRange': ['range start', 'range end'],
		#     'numTimeSteps': 'number of time steps',
		# }, {
		#     'sid': 'name of the scenario',
		#     'runStatus': 'finished',
		#     'timeRange': ['range start', 'range end'],
		#     'numTimeSteps': 'number of time steps'
		# }], 200

	def post(self, pid, model):
		'''
		Create a new scenario based on a reference
		:param pid:
		:param model:
		:return:
		'''
		# return the newly-created scenario together with CREATED 201 status code
		return {"a":model}, 201

class Input_List(Resource):

	def __init__(self):
		pass

	def get(self, format):
		if format == 'tree':
			weap_variables = weap_backend.get_WEAP_variables_tree('./model/WEAP_variables.json')
			leap_variables = leap_backend.get_LEAP_variables_tree('./model/LEAP_variables.json')
		if format == 'flat':
			weap_variables = weap_backend.get_WEAP_variables('./model/WEAP_variables.json')
			leap_variables = leap_backend.get_LEAP_variables('./model/LEAP_variables.json')
		data =  {"name": "FEW Nexus-Variables",
		        "children":[weap_variables[0],weap_variables[1], leap_variables[0], leap_variables[1]],
		       }
		print(data)
		return data , 200
	def post(self):

		return {"response": "POST method was just called!"}, 201
