from model import app
from flask import render_template, url_for, flash, redirect, request, abort, jsonify
from model.WEAPData import getFlowValue, Municipal
from flask_restful import reqparse, abort, Api, Resource
# from model.WEAPData import fff
import win32com.client
import win32com

# import pythoncom


api = Api(app)

# pythoncom.CoInitialize()
# WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
# value = WEAP.ResultValue(
# 	"\Supply and Resources\Transmission Links\\to Municipal\\from CAPWithdral:Flow[m^3]", 1986, 1,
# 	"5% Population Growth", 1986, 12, "Average")
flow = getFlowValue()


class Flow(Resource):
	def get(self):
		# flow = getFlowValue()
		return flow


class V1(Resource):
	def get(self):
		win32com.CoInitialize()
		WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
		value = WEAP.ResultValue(
			"\Supply and Resources\Transmission Links\\to Municipal\\from CAPWithdral:Flow[m^3]", 1986, 1,
			"5% Population Growth", 1986, 12, "Average")
		win32com.CoUninitialize()
		return value

class V2(Resource):
	def get(self):
		win32com.CoInitialize()
		WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
		value = WEAP.ResultValue(
			"\Supply and Resources\Transmission Links\\to Municipal\\from CAPWithdral:Flow[m^3]", 1986, 1,
			"5% Population Growth", 1986, 12, "Average")
		win32com.CoUninitialize()
		return value

# def delete(self, todo_id):
#     abort_if_todo_doesnt_exist(todo_id)
#     del TODOS[todo_id]
#     return '', 204
#
# def put(self, todo_id):
#     args = parser.parse_args()
#     task = {'task': args['task']}
#     TODOS[todo_id] = task
#     return task, 201


api.add_resource(Flow, "/flow")
api.add_resource(V1, "/v1")
api.add_resource(V2, "/v2")
