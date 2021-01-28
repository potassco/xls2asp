import unittest
import sys
import os
import clingo
import pandas as pd
import subprocess


class Context:
    def id(self, x):
        return x

    def seq(self, x, y):
        return [x, y]

# def sort_trace(trace):
#     return list(sorted(trace))

# def parse_model(m):
#     ret = []
#     for sym in m.symbols(shown=True):
#         if sym.name=="holds_map":
#             ret.append((sym.arguments[0].number, str(sym.arguments[1]) ))
#     return sort_trace(ret)

# def solve(const=[], files=[],inline_data=[]):
#     r = []
#     imax  = 20
#     ctl = clingo.Control(['0']+const, message_limit=0)
#     ctl.add("base", [], "")
#     for f in files:
#         ctl.load(f)
#     for d in inline_data:
#         ctl.add("base", [], d)
#     ctl.add("base",[],"#show holds_map/2.")
#     ctl.ground([("base", [])], context=Context())
#     ctl.solve(on_model= lambda m: r.append(parse_model(m)))
#     return sorted(r)

# def translate(constraint,file,extra=[]):
#     f = open("env/test/temporal_constraints/{}/{}".format(logic,file), "w")
#     f.write(constraint)
#     f.close()
#     command = 'make translate LOGIC={} CONSTRAINT={} APP=test INSTANCE=env/test/instances/empty.lp'.format(logic,file[:-3])
#     subprocess.check_output(command.split())

# def run_generate(constraint,mapping=None,horizon=3,file="formula_test.lp"):
#     translate(constraint,file)
#     files = ["outputs/test/{}/formula_test/empty/automaton.lp".format(logic),"./automata_run/run.lp","./automata_run/trace_generator.lp"]
#     if not mapping is None:
#         files.append(mapping)
#     return solve(["-c horizon={}".format(horizon)],files)

# def run_check(constraint,trace="",mapping="./env/test/glue.lp",encoding="",file="formula_test.lp",horizon=3,visualize=False):
#     translate(constraint,file)
#     if visualize:
#         command = "python scripts/viz.py {} {}".format(logic,file[:-3])
#         subprocess.check_output(command.split())

#     return solve(["-c horizon={}".format(horizon)],["outputs/test/{}/formula_test/empty/automaton.lp".format(logic),"./automata_run/run.lp",mapping],[trace,encoding])


class TestCase(unittest.TestCase):
    # clean up example
    def call_xls2asp(self, silent=False):
        command = 'python xls2asp.py --xls tests/tmp/data.xlsx --template ./tests/tmp/template.txt --output ./tests/tmp/output.lp'
        if silent:
            command_status = subprocess.call(
                command.split(), stderr=subprocess.DEVNULL)
        else:
            command_status = subprocess.call(command.split())
        return command_status

    def make_excel(self, data):
        df1 = pd.DataFrame(data)
        df1.to_excel("./tests/tmp/data.xlsx", index=False)

    def make_template(self, template):
        df1 = pd.DataFrame(template)
        df1.to_csv("./tests/tmp/template.txt", header=None,
                   index=None, sep=',', mode='w')

    def call_clingo(self, fact):
        ctl = clingo.Control()
        ctl.load("./tests/tmp/output.lp")
        ctl.add("base", [], ":- not "+fact+".")
        ctl.ground([("base", [])], context=Context())
        models = []
        ctl.solve(on_model=lambda m: models.append(m))
        return models

    def check_in_facts(self, fact):
        models = self.call_clingo(fact)
        self.assertGreater(
            len(models), 0, "Fact {} not in output".format(fact))


class TestMain(TestCase):

    def test_general(self):
        self.make_excel([['Dany', 'Hans', 20, 'male'], [
                        'Manuel', 'Vardi', 50, 'male']])
        self.make_template(
            [['Sheet1', 'row', 'string', 'string', 'int', 'constant']])
        self.assertEqual(self.call_xls2asp(), 0)
        self.check_in_facts('sheet1("Dany","Hans",20,male)')

    def test_time_iso(self):
        self.make_excel([['00:20:10'], [
                        '18:01:13']])
        self.make_template(
            [['Sheet1', 'row_indexed', 'time']])
        self.assertEqual(self.call_xls2asp(), 0)
        self.check_in_facts('sheet1(0,(0,20,10))')
        self.check_in_facts('sheet1(1,(18,1,13))')

        # Invalid syntax
        self.make_excel([['100:20:10'], [
                        '18:01:13']])
        self.make_template(
            [['Sheet1', 'row_indexed', 'time']])
        self.assertNotEqual(self.call_xls2asp(silent=True), 0)
