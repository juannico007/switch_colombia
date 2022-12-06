import glob, os
import pandas as pd

# find all the gen_cap files
search_pattern = os.path.join("outputs", "*", "gen_cap.csv")

frames = []
for path in glob.glob(search_pattern):
    # use the subdirectory name as the scenario
    parts = path.split(os.sep)
    scen = parts[-2].replace("_", " ")
    # read the output file and add to list
    df = pd.read_csv(path)
    df['scenario'] = scen
    frames.append(df)

# concatenate along vertical axis, replacing missing values with 0
df = pd.concat(frames, axis=0, sort=False)
# default group
df['technology'] = 'Menores'
# list of search key and technology labels
updates = [
    ('Hidro', 'Hidro'),
    ('Wind', 'Wind'),
    ('Auto-Cog', 'Auto-Cog'),
    ('Thermal', 'Thermal'),
    ('pv-solar', 'pv-solar'),
    ('Eolica', 'Eolica')
]
# df[df['gen_tech']=='Battery_Reg']
for search, tech in updates:
    df.loc[df['gen_tech'].str.contains(search), 'technology'] = tech

# aggregate GenCapacity values
agg = (
    df.groupby(['scenario', 'PERIOD', 'technology'])
    ['GenCapacity']
    .sum()
    .unstack(-1)   # create columns for each technology
    .loc[:, [              # reorder columns
        'Menores', 'Hidro',
        'Wind', 'Auto-Cog',
        'Thermal', 'pv-solar', 'Eolica'
    ]]
)

# add dummy rows for each scenario to make gaps when plotting later
agg = agg.reset_index()
dummies = pd.DataFrame({'scenario': agg['scenario'].unique()})
agg = agg.append(dummies, sort=False)
agg = agg.sort_values(['scenario', 'PERIOD'])

# remove scenario names except for first period,
# to improve Excel category labeling
agg.loc[agg['PERIOD'] != agg['PERIOD'].min(), 'scenario'] = ''

# change the dummy scenario names to ' ' to create gaps in Excel
agg.loc[agg['PERIOD'].isnull(), 'scenario'] = ' '

outfile = os.path.join('outputs', 'capacity_by_tech_by_scenario.csv')
agg.to_csv(outfile, index=False)
print("saved {}.".format(outfile))
