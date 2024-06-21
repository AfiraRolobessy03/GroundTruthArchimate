import pandas as pd

def notation_relation(notation):
    relationships = {
        'a': 'AccessRelationship',
        'c': 'CompositionRelationship',
        'f': 'FlowRelationship',
        'g': 'AggregationRelationship',
        'i': 'AssignmentRelationship',
        'n': 'InfluenceRelationship',
        'o': 'AssociationRelationship',
        'r': 'RealizationRelationship',
        's': 'SpecializationRelationship',
        't': 'TriggeringRelationship',
        'v': 'ServingRelationship',
    }

    return relationships.get(notation, 'Unknown notation')

def convert_notation_relation(notation):
    result = []
    
    for n in notation:
        result.append(notation_relation(n))

    return ",".join(result)
    
file_path = "Relasi.xlsx"
df = pd.read_excel(file_path, sheet_name='Application to Business', header=None)
data_sources = [
    ['Source', 'Relation', 'Target'],
]

for i in range(8, 17):
    source = df.iloc[i, 5]

    relation_row = 8
    for j in range(6, 19):
        target = df.iloc[7, j]
        relation = df.iloc[relation_row, j]
        normalized_relation = convert_notation_relation(relation)
        data_sources.append([source, normalized_relation, target])
        relation_row += 1

csv_df = pd.DataFrame(data_sources[1:], columns=data_sources[0])
csv_df.to_csv('application_to_business.csv', index=False)