"""
Limit the installation capacity per period.
"""
from __future__ import division
from pyomo.environ import *
import os

dependencies = 'switch_model.timescales', 'switch_model.balancing.load_zones',\
    'switch_model.financials', 'switch_model.energy_sources.properties.properties', \
    'switch_model.generators.core.build', 'switch_model.generators.core.dispatch'

def define_components(mod):

    mod.TECH_GEN_MW_RAW = Set(
    dimen=2,
    validate=lambda m, g, p: (g in m.GENERATION_PROJECTS) & (p in m.PERIODS))
    
    mod.TECH_GENS_CAP = Set(
        initialize=lambda m: set(g for (g, p) in m.TECH_GEN_MW_RAW))#,
        #doc="Dispatchable hydro projects")
    
    mod.limit_cap_mw = Param(
        mod.TECH_GEN_MW_RAW,
        within=NonNegativeReals,
        default=0.0)


    def Max_Build_Period_rule(m, g, p):

        cap_limit = 99999
        if g in m.TECH_GENS_CAP:
            cap_limit = m.limit_cap_mw[g, p]
        return (
            m.GenCapacity[g, p] <= cap_limit
        )
    mod.Max_Build_Period = Constraint(
        mod.GENERATION_PROJECTS, mod.PERIODS,
        rule=Max_Build_Period_rule)


def load_inputs(mod, switch_data, inputs_dir):
    """

    Import limits data.

    generation_limits_build.csv
        GENERATION_PROJECT, INVESTMENT_PERIOD, limit_cap_mw

    """
    switch_data.load_aug(
        # optional=True,
        filename=os.path.join(inputs_dir, 'generation_limits_build.csv'),
        autoselect=True,
        index=mod.TECH_GEN_MW_RAW,
        param=(mod.limit_cap_mw)
    )