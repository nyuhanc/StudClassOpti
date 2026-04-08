# Approach using Constraint Programming (v2 - improved)
#
# Changes vs main_CP.py:
# - BUG FIX: Objective function nat_sci terms now use nat_sci priorities (1-3),
#   not language priorities (1-5). Previously the nat_sci scoring was entirely wrong.
# - BUG FIX: Constraints 11 & 12 now use Python-level conditionals (since
#   _nat_sci_priorities are Python ints, not CP-SAT variables). The original
#   AddBoolOr([python_bool, ...]) calls silently misbehaved.
# - BUG FIX: Constraint 14 now uses a Python-level conditional instead of trying
#   to constrain a BoolVar via model.Add(python_int == python_int).
# - IMPROVEMENT: Solver runs with parallel workers and a per-shuffle time limit.
# - IMPROVEMENT: Number of shuffles increased from 10 to 20.
# - IMPROVEMENT: nat_sci top-priority penalty now correctly fires once (not twice)
#   when neither slot contains the student's top nat_sci choice.

from ortools.sat.python import cp_model
import pandas as pd
import itertools
import os
import numpy as np

# Metainfo
# - Student: name of the student
# - Language (German, Spanish, Italian, Russian, French): each student must be assigned a single language
#   based on the priority that they have set to the languages: 1 - highest priority, 5 - lowest priority
# - Natural Science (Biology, Physics, Chemistry): each student must be assigned two natural science
#   classes based on the priority that they have set to the classes: 1 - highest priority, 3 - lowest priority
# - Schoolmate: the name of the schoolmate that the student wants to be in the same class with
# - NationalTestScore: the score of the student in the national test

# Global variables
num_of_classes = 3
max_class_size = 28

# Solver tuning
shuffles = 20
time_limit_per_shuffle = 120  # seconds
num_workers = 8

# ---------------------------------------------------

# Load and preprocess the data

filename = 'students_list_2025'
appendix = '.xlsx'

data = pd.read_excel(filename + appendix, engine='openpyxl')

if len(data) > 84:
    data = data.head(84)

data['Chemistry'] = data['Chemistry'].astype(int)
data['Student'] = data['Student'].astype(int)

students = data['Student'].tolist()

languages = [
    'French',    # -> language 1
    'Italian',   # -> language 2
    'German',    # -> language 3
    'Russian',   # -> language 4
    'Spanish'    # -> language 5
]
nat_sci_classes = [
    'Biology',   # -> natural science 1
    'Physics',   # -> natural science 2
    'Chemistry'  # -> natural science 3
]

print('Errors in the data:')

for student in students:
    _langs_pris = []
    for lang in languages:
        _langs_pris.append(data[data['Student'] == student][lang].values[0])
    if set(_langs_pris) != set([1, 2, 3, 4, 5]):
        print(f"Wrong values for language priorities for student {student}")

input('Continue (press enter) only if there are no errors in the data. Otherwise, fix the errors and run the script again.')

# Count the number of students that share the same top-2 natural science preferences
best_ns_match_pair = (None, None)
best_ns_match = 0
for ns1, ns2 in [tup for tup in list(itertools.product(nat_sci_classes, repeat=2)) if tup[0] < tup[1]]:

    ns_match = 0

    for student in students:
        if data[data['Student'] == student][ns1].values[0] == 1 and data[data['Student'] == student][ns2].values[0] == 2:
            ns_match += 1
        elif data[data['Student'] == student][ns2].values[0] == 1 and data[data['Student'] == student][ns1].values[0] == 2:
            ns_match += 1

    if ns_match > best_ns_match:
        best_ns_match = ns_match
        best_ns_match_pair = (ns1, ns2)

    print(f'Number of students with the same natural science classes {ns1} and {ns2}: {ns_match}')

print(f'The best match is between {best_ns_match_pair[0]} and {best_ns_match_pair[1]} with {best_ns_match} students having the same preferences.')
print('With respect to the above information, please adjust the constraint(s) 6a (and 6b if exists) in the model accordingly.')
input('Press enter to continue.')

# ---------------------------------------------------

best_score = 0
best_data = None

for shuffle_idx in range(shuffles):
    print(f'\n--- Shuffle {shuffle_idx + 1}/{shuffles} ---')

    data = data.sample(frac=1).reset_index(drop=True)
    students = data['Student'].tolist()

    model = cp_model.CpModel()

    # Create decision variables
    var_dict = {}
    for student in students:
        var_dict[str(student)] = student
        var_dict[str(student) + '_class'] = model.NewIntVar(1, num_of_classes, str(student) + '_class')
        var_dict[str(student) + '_lang'] = model.NewIntVar(1, 5, str(student) + '_lang')
        var_dict[str(student) + '_nat_sci_1'] = model.NewIntVar(lb=1, ub=3, name=str(student) + '_nat_sci_1')
        var_dict[str(student) + '_nat_sci_2'] = model.NewIntVar(lb=1, ub=3, name=str(student) + '_nat_sci_2')

    # Build preference priority lists (Python ints, not CP-SAT vars)
    for student in students:
        lang_ratings = []
        for lang in languages:
            lang_ratings.append(data[data['Student'] == student][lang].values[0])
        # _lang_priorities[i] = language index (1-5) that has rank (i+1)
        var_dict[str(student) + '_lang_priorities'] = []
        for i in range(1, 6):
            var_dict[str(student) + '_lang_priorities'].append(lang_ratings.index(i) + 1)

        nat_sci_ratings = []
        for nat_sci in nat_sci_classes:
            nat_sci_ratings.append(data[data['Student'] == student][nat_sci].values[0])
        # _nat_sci_priorities[i] = nat_sci index (1-3) that has rank (i+1)
        var_dict[str(student) + '_nat_sci_priorities'] = []
        for i in range(1, 4):
            var_dict[str(student) + '_nat_sci_priorities'].append(nat_sci_ratings.index(i) + 1)

    # ---------- Constraints ----------
    cons_names = []

    cons_names.append('1. Maximally max_class_size students can be assigned to each class')
    for i in range(1, num_of_classes + 1):
        students_in_class_i = []
        for student in students:
            is_in_class_i = model.NewBoolVar(f'{student}_in_class_{i}')
            model.Add(var_dict[str(student) + '_class'] == i).OnlyEnforceIf(is_in_class_i)
            model.Add(var_dict[str(student) + '_class'] != i).OnlyEnforceIf(is_in_class_i.Not())
            students_in_class_i.append(is_in_class_i)
        model.Add(sum(students_in_class_i) <= max_class_size)

    cons_names.append('2. Maximally 1*max_class_size students can have the same language')
    for i in range(1, 6):
        students_with_lang_i = []
        for student in students:
            has_lang_i = model.NewBoolVar(f'{student}_has_lang_{i}')
            model.Add(var_dict[str(student) + '_lang'] == i).OnlyEnforceIf(has_lang_i)
            model.Add(var_dict[str(student) + '_lang'] != i).OnlyEnforceIf(has_lang_i.Not())
            students_with_lang_i.append(has_lang_i)
        model.Add(sum(students_with_lang_i) <= 1 * max_class_size)

    cons_names.append('3. Maximally 3 * max_class_size students can have the same natural science classes')
    for i in range(1, 4):
        students_with_nat_sci = []
        for student in students:
            has_nat_sci_1 = model.NewBoolVar(f'{student}_has_nat_sci_1_{i}')
            model.Add(var_dict[str(student) + '_nat_sci_1'] == i).OnlyEnforceIf(has_nat_sci_1)
            model.Add(var_dict[str(student) + '_nat_sci_1'] != i).OnlyEnforceIf(has_nat_sci_1.Not())

            has_nat_sci_2 = model.NewBoolVar(f'{student}_has_nat_sci_2_{i}')
            model.Add(var_dict[str(student) + '_nat_sci_2'] == i).OnlyEnforceIf(has_nat_sci_2)
            model.Add(var_dict[str(student) + '_nat_sci_2'] != i).OnlyEnforceIf(has_nat_sci_2.Not())

            students_with_nat_sci.append(has_nat_sci_1)
            students_with_nat_sci.append(has_nat_sci_2)

        model.Add(sum(students_with_nat_sci) <= max_class_size * 3)

    cons_names.append('4. Natural classes 1 and 2 of the same student must be different')
    for student in students:
        model.Add(var_dict[str(student) + '_nat_sci_1'] != var_dict[str(student) + '_nat_sci_2'])

    # cons_names.append('5. Students that want to be in the same class must be in the same class')
    # for student in students:
    #     schoolmate = data[data['Student'] == student]['Schoolmate'].values[0]
    #     if schoolmate != 'None' and schoolmate != 0:
    #         model.Add(var_dict[str(student) + '_class'] == var_dict[str(schoolmate) + '_class'])

    cons_names.append('6. Students in one unspecified class are assigned to the most preferred natural science classes')
    ns1_index = nat_sci_classes.index(best_ns_match_pair[0]) + 1
    ns2_index = nat_sci_classes.index(best_ns_match_pair[1]) + 1

    best_matching_class = model.NewIntVar(1, num_of_classes, 'best_matching_class')

    for student in students:
        is_assigned_ns1_ns2 = model.NewBoolVar(f'{student}_is_assigned_ns1_ns2')
        is_assigned_ns2_ns1 = model.NewBoolVar(f'{student}_is_assigned_ns2_ns1')

        model.Add(var_dict[str(student) + '_nat_sci_1'] == ns1_index).OnlyEnforceIf(is_assigned_ns1_ns2)
        model.Add(var_dict[str(student) + '_nat_sci_2'] == ns2_index).OnlyEnforceIf(is_assigned_ns1_ns2)
        model.Add(var_dict[str(student) + '_nat_sci_1'] == ns2_index).OnlyEnforceIf(is_assigned_ns2_ns1)
        model.Add(var_dict[str(student) + '_nat_sci_2'] == ns1_index).OnlyEnforceIf(is_assigned_ns2_ns1)

        is_in_best_matching_class = model.NewBoolVar(f'{student}_in_best_matching_class')
        model.Add(var_dict[str(student) + '_class'] == best_matching_class).OnlyEnforceIf(is_in_best_matching_class)
        model.Add(var_dict[str(student) + '_class'] != best_matching_class).OnlyEnforceIf(is_in_best_matching_class.Not())

        model.Add(is_assigned_ns1_ns2 + is_assigned_ns2_ns1 == is_in_best_matching_class)

    cons_names.append(
        '8. Male students are split into exactly two classes, with each containing 4-7 males and one class containing none')
    male_students = data[data['Gender'] == 'm']['Student'].tolist()
    if male_students:
        male_counts_per_class = []
        for i in range(1, num_of_classes + 1):
            male_in_class_i = []
            for male_student in male_students:
                is_in_class_i = model.NewBoolVar(f'{male_student}_male_in_class_{i}')
                model.Add(var_dict[str(male_student) + '_class'] == i).OnlyEnforceIf(is_in_class_i)
                model.Add(var_dict[str(male_student) + '_class'] != i).OnlyEnforceIf(is_in_class_i.Not())
                male_in_class_i.append(is_in_class_i)
            male_count_in_class_i = model.NewIntVar(0, len(male_students), f'male_count_in_class_{i}')
            model.Add(male_count_in_class_i == sum(male_in_class_i))
            male_counts_per_class.append(male_count_in_class_i)

        classes_with_males = []
        for i in range(num_of_classes):
            has_males = model.NewBoolVar(f'class_{i + 1}_has_males')
            model.Add(male_counts_per_class[i] >= 1).OnlyEnforceIf(has_males)
            model.Add(male_counts_per_class[i] == 0).OnlyEnforceIf(has_males.Not())
            classes_with_males.append(has_males)

        model.Add(sum(classes_with_males) == 2)

        for count in male_counts_per_class:
            is_nonzero = model.NewBoolVar(f'{count.Name()}_nonzero')
            model.Add(count >= 1).OnlyEnforceIf(is_nonzero)
            model.Add(count == 0).OnlyEnforceIf(is_nonzero.Not())
            model.Add(count >= 4).OnlyEnforceIf(is_nonzero)
            model.Add(count <= 7).OnlyEnforceIf(is_nonzero)

    cons_names.append('9. All students who chose language 4 as their first priority must be assigned to language 4')
    lang_4 = 4
    for student in students:
        # _lang_priorities[0] is a Python int — use a Python-level conditional
        if var_dict[str(student) + '_lang_priorities'][0] == lang_4:
            model.Add(var_dict[str(student) + '_lang'] == lang_4)

    cons_names.append('10. There can be only max_class_size students assigned to the natural science class 2 (physics)')
    specific_ns_index = 2
    students_assigned_to_ns2 = []
    for student in students:
        assigned_ns2_as_ns1 = model.NewBoolVar(f'{student}_assigned_ns2_as_ns1')
        assigned_ns2_as_ns2 = model.NewBoolVar(f'{student}_assigned_ns2_as_ns2')

        model.Add(var_dict[str(student) + '_nat_sci_1'] == specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns1)
        model.Add(var_dict[str(student) + '_nat_sci_1'] != specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns1.Not())
        model.Add(var_dict[str(student) + '_nat_sci_2'] == specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns2)
        model.Add(var_dict[str(student) + '_nat_sci_2'] != specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns2.Not())

        students_assigned_to_ns2.append(assigned_ns2_as_ns1)
        students_assigned_to_ns2.append(assigned_ns2_as_ns2)

    model.Add(sum(students_assigned_to_ns2) <= max_class_size)

    # FIX: _nat_sci_priorities are Python ints — use Python-level conditionals
    cons_names.append('11. All students with Physics as first or second priority must be assigned to Physics')
    physics_index = nat_sci_classes.index('Physics') + 1
    for student in students:
        first_priority_ns = var_dict[str(student) + '_nat_sci_priorities'][0]
        second_priority_ns = var_dict[str(student) + '_nat_sci_priorities'][1]

        if first_priority_ns == physics_index or second_priority_ns == physics_index:
            has_physics_as_ns1 = model.NewBoolVar(f'{student}_assigned_physics_ns1')
            has_physics_as_ns2 = model.NewBoolVar(f'{student}_assigned_physics_ns2')

            model.Add(var_dict[str(student) + '_nat_sci_1'] == physics_index).OnlyEnforceIf(has_physics_as_ns1)
            model.Add(var_dict[str(student) + '_nat_sci_1'] != physics_index).OnlyEnforceIf(has_physics_as_ns1.Not())
            model.Add(var_dict[str(student) + '_nat_sci_2'] == physics_index).OnlyEnforceIf(has_physics_as_ns2)
            model.Add(var_dict[str(student) + '_nat_sci_2'] != physics_index).OnlyEnforceIf(has_physics_as_ns2.Not())

            model.AddBoolOr([has_physics_as_ns1, has_physics_as_ns2])

    # FIX: same — use Python-level conditional
    cons_names.append('12. Students with Physics as 3rd priority must not be assigned to Physics')
    for student in students:
        third_priority_ns = var_dict[str(student) + '_nat_sci_priorities'][2]
        if third_priority_ns == physics_index:
            model.Add(var_dict[str(student) + '_nat_sci_1'] != physics_index)
            model.Add(var_dict[str(student) + '_nat_sci_2'] != physics_index)

    cons_names.append('13. Any student assigned Physics must also be assigned Biology in the other slot')
    biology_index = nat_sci_classes.index('Biology') + 1
    physics_index = nat_sci_classes.index('Physics') + 1

    for student in students:
        ns1 = var_dict[str(student) + '_nat_sci_1']
        ns2 = var_dict[str(student) + '_nat_sci_2']

        ns1_is_physics = model.NewBoolVar(f'{student}_ns1_is_physics')
        ns2_is_physics = model.NewBoolVar(f'{student}_ns2_is_physics')

        model.Add(ns1 == physics_index).OnlyEnforceIf(ns1_is_physics)
        model.Add(ns1 != physics_index).OnlyEnforceIf(ns1_is_physics.Not())
        model.Add(ns2 == physics_index).OnlyEnforceIf(ns2_is_physics)
        model.Add(ns2 != physics_index).OnlyEnforceIf(ns2_is_physics.Not())

        model.Add(ns2 == biology_index).OnlyEnforceIf(ns1_is_physics)
        model.Add(ns1 == biology_index).OnlyEnforceIf(ns2_is_physics)

    # FIX: _lang_priorities[0] is a Python int — use Python-level conditional
    cons_names.append('14. Students with Spanish as top priority must not be assigned Italian')
    spanish_index = languages.index('Spanish') + 1
    italian_index = languages.index('Italian') + 1
    for student in students:
        top_priority_lang = var_dict[str(student) + '_lang_priorities'][0]
        if top_priority_lang == spanish_index:
            model.Add(var_dict[str(student) + '_lang'] != italian_index)

    # ---------- Objective function ----------

    lang_importance = 1
    lang_penalty = 10
    nat_sci_1_importance = 100
    nat_sci_2_importance = 1
    nat_sci_penalty = 100
    stratification = 2

    objective_terms = []
    for student in students:

        chosen_lang = var_dict[str(student) + '_lang']
        chosen_nat_sci_1 = var_dict[str(student) + '_nat_sci_1']
        chosen_nat_sci_2 = var_dict[str(student) + '_nat_sci_2']

        # --- Language scoring ---
        for i in range(1, 6):
            lang_at_priority_i = var_dict[str(student) + '_lang_priorities'][i - 1]

            assigned_lang_i = model.NewBoolVar(f'{student}_chosen_lang_{i}')
            model.Add(chosen_lang == lang_at_priority_i).OnlyEnforceIf(assigned_lang_i)
            model.Add(chosen_lang != lang_at_priority_i).OnlyEnforceIf(assigned_lang_i.Not())

            objective_terms.append(lang_importance * assigned_lang_i * (5 - i) ** stratification)

            if i == 1:
                objective_terms.append(-lang_penalty * assigned_lang_i.Not())

        # --- Natural science scoring ---
        # FIX: use _nat_sci_priorities (range 1-3), not _lang_priorities (range 1-5)
        nat_sci_1_assigned_top = None
        nat_sci_2_assigned_top = None

        for i in range(1, 4):
            nat_sci_at_priority_i = var_dict[str(student) + '_nat_sci_priorities'][i - 1]

            assigned_nat_sci_1_i = model.NewBoolVar(f'{student}_chosen_nat_sci_1_{i}')
            model.Add(chosen_nat_sci_1 == nat_sci_at_priority_i).OnlyEnforceIf(assigned_nat_sci_1_i)
            model.Add(chosen_nat_sci_1 != nat_sci_at_priority_i).OnlyEnforceIf(assigned_nat_sci_1_i.Not())

            objective_terms.append(nat_sci_1_importance * assigned_nat_sci_1_i * (3 - i) ** stratification)

            assigned_nat_sci_2_i = model.NewBoolVar(f'{student}_chosen_nat_sci_2_{i}')
            model.Add(chosen_nat_sci_2 == nat_sci_at_priority_i).OnlyEnforceIf(assigned_nat_sci_2_i)
            model.Add(chosen_nat_sci_2 != nat_sci_at_priority_i).OnlyEnforceIf(assigned_nat_sci_2_i.Not())

            objective_terms.append(nat_sci_2_importance * assigned_nat_sci_2_i * (3 - i) ** stratification)

            if i == 1:
                nat_sci_1_assigned_top = assigned_nat_sci_1_i
                nat_sci_2_assigned_top = assigned_nat_sci_2_i

        # Penalize once if neither slot contains the student's top nat_sci choice
        top_ns_in_any_slot = model.NewBoolVar(f'{student}_top_ns_in_any_slot')
        model.AddBoolOr([nat_sci_1_assigned_top, nat_sci_2_assigned_top]).OnlyEnforceIf(top_ns_in_any_slot)
        model.AddBoolAnd([nat_sci_1_assigned_top.Not(), nat_sci_2_assigned_top.Not()]).OnlyEnforceIf(top_ns_in_any_slot.Not())
        objective_terms.append(-nat_sci_penalty * top_ns_in_any_slot.Not())

    model.Maximize(sum(objective_terms))

    # Solve with parallel workers and a time limit
    solver = cp_model.CpSolver()
    solver.parameters.num_search_workers = num_workers
    solver.parameters.max_time_in_seconds = time_limit_per_shuffle

    status = solver.Solve(model)

    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        score = solver.ObjectiveValue()
        print(f'Score: {score}  (best so far: {max(best_score, score)})')

        if score > best_score:
            best_score = score
            best_data = data.copy()

            best_data['Class'] = None
            best_data['Language'] = None
            best_data['NatSci1'] = None
            best_data['NatSci2'] = None

            for student in students:
                best_data.loc[data['Student'] == student, 'Class'] = solver.Value(var_dict[str(student) + '_class'])
                best_data.loc[data['Student'] == student, 'Language'] = languages[solver.Value(var_dict[str(student) + '_lang']) - 1]
                best_data.loc[data['Student'] == student, 'NatSci1'] = nat_sci_classes[solver.Value(var_dict[str(student) + '_nat_sci_1']) - 1]
                best_data.loc[data['Student'] == student, 'NatSci2'] = nat_sci_classes[solver.Value(var_dict[str(student) + '_nat_sci_2']) - 1]
    else:
        print(f'Shuffle {shuffle_idx + 1}: No solution found.')

if best_data is None:
    print('No solution found in any shuffle. Exiting.')
    exit(1)

best_data = best_data.sort_values(by=['Class', 'Language', 'NatSci1', 'NatSci2'])

print(f'\nBest total score: {best_score}')

class_sizes = best_data['Class'].value_counts().sort_index()
print('\nClass sizes:')
print(class_sizes)

results_name = input('Save into dir name (inside results dir): ')
if not os.path.exists('results/' + results_name):
    os.makedirs('results/' + results_name)

best_data.to_excel('results/' + results_name + '/' + filename + '_' + results_name + appendix, index=False)

with open('results/' + results_name + '/' + filename + '_' + results_name + '_model_parameters.txt', 'w') as f:
    f.write('Objective function parameters:\n')
    f.write(f'lang_importance = {lang_importance}\n')
    f.write(f'lang_penalty = {lang_penalty}\n')
    f.write(f'nat_sci_1_importance = {nat_sci_1_importance}\n')
    f.write(f'nat_sci_2_importance = {nat_sci_2_importance}\n')
    f.write(f'nat_sci_penalty = {nat_sci_penalty}\n')
    f.write(f'stratification = {stratification}\n')
    f.write('\n')
    f.write('Constraints:\n')
    for i, cons in enumerate(cons_names):
        f.write(f'{i + 1}. {cons}\n')
    f.write('\n')
    f.write('Other information:\n')
    f.write(f'Max class size: {max_class_size}\n')
    f.write(f'Number of classes: {num_of_classes}\n')
    f.write(f'Number of shuffles: {shuffles}\n')
    f.write(f'Time limit per shuffle: {time_limit_per_shuffle}s\n')
    f.write(f'Search workers: {num_workers}\n')
    f.write(f'Best score: {best_score}\n')
    f.write(f'Best match pair of natural science classes: {best_ns_match_pair[0]} and {best_ns_match_pair[1]}')
