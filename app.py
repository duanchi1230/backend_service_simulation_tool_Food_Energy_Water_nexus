"""
RUN this script to start the FEWSim backend
"""

#########!env python3
from flask import Flask, request
from flask_restful import Resource, Api, abort
import sys, json

from resources.project import Project, ProjectList
from resources.scenario import Scenario, ScenarioList, Input_List, Run_Sceanrios, Get_Run_Log, Load_Existing_Scenarios, \
    Get_Coupled_Variable, Load_Sustainability_Index, Save_Simulation_Result, Sensitivity_Graph, Load_Simulation_History


# route the RESTFUL API addresses
app = Flask(__name__)
api = Api(app)
api.add_resource(ProjectList, '/proj')
api.add_resource(Project, '/proj/<string:pid>')
api.add_resource(ScenarioList, '/proj/<string:pid>/<string:model>/scenario')
api.add_resource(Scenario, '/proj/<string:pid>/<string:model>/scenario/<string:sid>')
api.add_resource(Input_List, '/inputs/<string:format>')
api.add_resource(Run_Sceanrios, '/run/<string:scenario>')
api.add_resource(Get_Run_Log, '/log')
api.add_resource(Load_Existing_Scenarios, '/load-scenarios')
api.add_resource(Get_Coupled_Variable, '/get-coupled-parameters')
api.add_resource(Load_Sustainability_Index, '/get-sustainability-index')
api.add_resource(Sensitivity_Graph, '/sensitivity-graph')
api.add_resource(Load_Simulation_History, '/load-simulation-history')

if __name__ == '__main__':

    if len(sys.argv) < 2:
        port = 5000
    else:
        port = int(sys.argv[1])

    app.run(host='0.0.0.0', port=port)

if __name__ == "__main__":
    app.run(debug=True)

# with open('./model/LEAP_variables.json') as f:
# 	variables = json.load(f)
# 	print(variables)
