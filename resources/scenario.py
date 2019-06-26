from flask import Flask, request
from flask_restful import Resource, Api, abort


class Scenario(Resource):

    def __init__(self):
        pass


    def get(self, pid, model, sid):

        # Chi: implement a function that gets a scenario based on the "model" and the "sid"
        # The returned dict includes all the input and output variables

        return {
            'sid': 'name of the scenario',
            'runStatus': 'finished',
            'timeRange': ['range start', 'range end'],
            'numTimeSteps': 'number of time steps',
            'var': [{
                'name': 'var1',
                'type': 'input or output',
                'format': 'number or series',
                'value': 'the actual value'
            }]
        }, 200


class ScenarioList(Resource):

    def __init__(self):
        pass


    def get(self, pid, model):

        # Chi: implement a function that gets all the existing scenarios
        # and return their brief summary as follows:

        return [{
            'sid': 'name of the scenario',
            'runStatus': 'finished',
            'timeRange': ['range start', 'range end'],
            'numTimeSteps': 'number of time steps'
        }, {
            'sid': 'name of the scenario',
            'runStatus': 'finished',
            'timeRange': ['range start', 'range end'],
            'numTimeSteps': 'number of time steps'
        }], 200


    def post(self, pid, model):
        """
        Create a new scenario based on a reference
        :param pid:
        :param model:
        :return:
        """

        # return the newly-created scenario together with CREATED 201 status code
        return {}, 201