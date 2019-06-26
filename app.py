#!env python3

from flask import Flask, request
from flask_restful import Resource, Api, abort
import sys

from resources.project import Project, ProjectList
from resources.scenario import Scenario, ScenarioList

app = Flask(__name__)
api = Api(app)

api.add_resource(ProjectList, '/proj')
api.add_resource(Project, '/proj/<string:pid>')
api.add_resource(ScenarioList, '/proj/<string:pid>/<string:model>/scenario')
api.add_resource(Scenario, '/proj/<string:pid>/<string:model>/scenario/<string:sid>')


if __name__ == '__main__':
    if len(sys.argv) < 2:
        port = 5000
    else:
        port = int(sys.argv[1])

    app.run(host='0.0.0.0', port=port)
