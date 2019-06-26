from flask import Flask, request
from flask_restful import Resource, Api, abort

DUMMY_UUID = 'f69e000c-6baa-4802-b907-699d75d6fad3'
DUMMY_PROJ = {
    'pid': DUMMY_UUID,
    'createTime': '2019-06-26 09:26:03.478039',
    'supportedModels': 'weap'
}


class Project(Resource):

    def __init__(self):
        pass

    def get(self, pid):
        """
        Dummy filling of only one project
        :return:
        """

        return DUMMY_PROJ


class ProjectList(Resource):
    def __init__(self):
        pass

    def get(self):
        """
        Listing all existing projects
        :return:
        """
        return [DUMMY_PROJ]
