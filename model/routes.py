# from model import app
# from model.WEAPData import get_WEAP_FlowValue, get_WEAP_para_value
# from flask_restful import Api, Resource
#
#
#
# # import pythoncom
#
#
# api = Api(app)
#
# # pythoncom.CoInitialize()
# # WEAP = win32com.client.Dispatch("WEAP.WEAPApplication")
# # value = WEAP.ResultValue(
# # 	"\Supply and Resources\Transmission Links\\to Municipal\\from CAPWithdral:Flow[m^3]", 1986, 1,
# # 	"5% Population Growth", 1986, 12, "Average")
# # flow = get_WEAP_FlowValue()
#
#
# class Flow(Resource):
# 	def get(self):
# 		# flow = getFlowValue()
# 		return flow
#
#
# class V1(Resource):
# 	def get(self):
# 		path = {"branch": "\Demand Sites\Municipal", "variable": "Annual Activity Level"}
# 		para = get_WEAP_para_value(path)
# 		return para
#
# class V2(Resource):
# 	def get(self):
# 		path = {"branch": "\Demand Sites\Municipal", "variable": "Annual Activity Level"}
# 		para = get_WEAP_para_value(path)
#
# 		return para
#
# # def delete(self, todo_id):
# #     abort_if_todo_doesnt_exist(todo_id)
# #     del TODOS[todo_id]
# #     return '', 204
# #
# # def put(self, todo_id):
# #     args = parser.parse_args()
# #     task = {'task': args['task']}
# #     TODOS[todo_id] = task
# #     return task, 201
#
# #
# # api.add_resource(Flow, "weap/flow")
# # api.add_resource(V1, "/v1")
# # api.add_resource(V2, "/v2")
