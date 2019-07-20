from flask import Flask, request
from flask_restful import Resource, Api, abort
# from model import WEAP_Data as WEAP_model
from model import WEAP_Simple_Model as WEAP_model

class Scenario(Resource):

	def __init__(self):
		pass

	def get(self, pid, model, sid):
		sid = int(sid)

		# Chi: implement a function that gets a scenario based on the 'model' and the 'sid'
		# The returned dict includes all the input and output variables
		scenarios = ['Reference', '5% Population Growth', '10% Population Growth']
		path = {'branch': '\Demand Sites\Municipal', 'variable': 'Annual Activity Level'}

		# para = WEAP_Simple_Model.get_WEAP_para_value(path)
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
			path = {'branch': '\Demand Sites\Municipal', 'variable': 'Annual Activity Level'}
			year = [1986, 2008]
			# para = WEAP_Data.get_WEAP_para_value(path)
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
		return {}, 201
