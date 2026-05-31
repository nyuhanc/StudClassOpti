# Approach using Constraint Programming

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
#   classes based on teh priority that they have set to the classes: 1 - highest priority, 5 - lowest priority
# - Schoolmate: the name of the schoolmate that the student wants to be in the same class with
# - NationalTestScore: the score of the student in the national test

# Global variables
num_of_classes = 3
max_class_size = 28

# ---------------------------------------------------

# Load and preprocess the data

# Import the data
filename = 'students_list_2025'
appendix = '.xlsx'

data = pd.read_excel(filename + appendix, engine='openpyxl')

# --- CORRECTIONS EXAMPLE---

# # Convert Schoolmate values that are nan to 0
# data['Schoolmate'] = data['Schoolmate'].fillna(0)
#
# # Convert Schoolmate values to integers
# data['Schoolmate'] = data['Schoolmate'].astype(int)

# If numbers of rows in data greater than 84 retain only the first 84 columns
if len(data) > 84:
    data = data.head(84)

# Convert Values in Chemistry to int64
data['Chemistry'] = data['Chemistry'].astype(int)

# Convert student names to int64
data['Student'] = data['Student'].astype(int)

# # Save the corrected data to a new Excel file
# data.to_excel(filename + appendix, index=False)

# --- END OF CORRECTIONS ---

# Get the list of students and convert them to integers
students = data['Student'].tolist()

# Marking classes with numbers
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

# Check for missing or wrong values
for student in students:

    _langs_pris = []
    for lang in languages:
        _langs_pris.append(data[data['Student'] == student][lang].values[0])
    if set(_langs_pris) != set([1,2,3,4,5]):
        print(f"Wrong values for language priorities for student {student}")


input('Continue (press enter) only if there are no errors in the data. Otherwise, fix the errors and run the script again.')

# Count the number of students that have same preferred natural science classes 1 and 2 (and 2 and 1)
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

shuffles = 10
best_score = 0
for _ in range(shuffles):

    # Shuffle the data
    data = data.sample(frac=1).reset_index(drop=True)

    students = data['Student'].tolist()

    # Create the model
    model = cp_model.CpModel()

    # Create the variables
    var_dict = {}
    for student in students:
        var_dict[str(student)] = student
        var_dict[str(student) + '_class'] = model.NewIntVar(1, num_of_classes, str(student) + '_class')
        var_dict[str(student) + '_lang'] = model.NewIntVar(1, 5, str(student) + '_lang')
        var_dict[str(student) + '_nat_sci_1'] = model.NewIntVar(lb=1, ub=3, name=str(student) + '_nat_sci_1')
        var_dict[str(student) + '_nat_sci_2'] = model.NewIntVar(lb=1, ub=3, name=str(student) + '_nat_sci_2')

    # Add information on the language preferences to the students_dict
    # in the form of lists of priorities as [top_priority_language_index,
    # ..., lowest_priority_language_index]. Same for natural sciences.
    for student in students:

        lang_priorities = []
        for lang in languages:
            lang_priorities.append(data[data['Student'] == student][lang].values[0])
        var_dict[str(student) + '_lang_priorities'] = []
        for i in range(1, 6):
            var_dict[str(student) + '_lang_priorities'].append(lang_priorities.index(i) + 1)

        nat_sci_priorities = []
        for nat_sci in nat_sci_classes:
            nat_sci_priorities.append(data[data['Student'] == student][nat_sci].values[0])
        var_dict[str(student) + '_nat_sci_priorities'] = []
        for i in range(1, 4):
            var_dict[str(student) + '_nat_sci_priorities'].append(nat_sci_priorities.index(i) + 1)

    # ---------- Create the constraints -------------------
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

            # Create boolean variables that are 1 if the student has the natural science class i
            has_nat_sci_1 = model.NewBoolVar(f'{student}_has_nat_sci_1_{i}')
            model.Add(var_dict[str(student) + '_nat_sci_1'] == i).OnlyEnforceIf(has_nat_sci_1)
            model.Add(var_dict[str(student) + '_nat_sci_1'] != i).OnlyEnforceIf(has_nat_sci_1.Not())

            # Create boolean variables that are 1 if the student has the natural science class i
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

    # cons_names.append('6. Every student in class 1 is assigned to natural science classes that are the most preferred by the students')
    # ns1, ns2 = best_ns_match_pair
    # ns1_index = nat_sci_classes.index(ns1) + 1  # Convert class name to index
    # ns2_index = nat_sci_classes.index(ns2) + 1  # Convert class name to index
    #
    # for student in students:
    #     # Create auxiliary boolean variables that are 1 if the student is assigned to the natural science classes ns1 and ns2 in any order
    #     is_assigned_ns1_ns2 = model.NewBoolVar(f'{student}_is_assigned_ns1_ns2')
    #     is_assigned_ns2_ns1 = model.NewBoolVar(f'{student}_is_assigned_ns2_ns1')
    #
    #     # Ensure the natural science classes match the best match pair in any order
    #     model.Add(var_dict[str(student) + '_nat_sci_1'] == ns1_index).OnlyEnforceIf(is_assigned_ns1_ns2)
    #     model.Add(var_dict[str(student) + '_nat_sci_2'] == ns2_index).OnlyEnforceIf(is_assigned_ns1_ns2)
    #     model.Add(var_dict[str(student) + '_nat_sci_1'] == ns2_index).OnlyEnforceIf(is_assigned_ns2_ns1)
    #     model.Add(var_dict[str(student) + '_nat_sci_2'] == ns1_index).OnlyEnforceIf(is_assigned_ns2_ns1)
    #
    #     # Ensure that if a student is in class 1, they must be assigned to the best match pair of natural science classes in any order
    #     is_in_class_1 = model.NewBoolVar(f'{student}_in_class_1')
    #     model.Add(var_dict[str(student) + '_class'] == 1).OnlyEnforceIf(is_in_class_1)
    #     model.Add(var_dict[str(student) + '_class'] != 1).OnlyEnforceIf(is_in_class_1.Not())
    #
    #     # This line ensures that if the student is in class 1, they must satisfy one of the conditions
    #     model.Add(is_assigned_ns1_ns2 + is_assigned_ns2_ns1 == is_in_class_1)

    cons_names.append('6. Students in one unspecified class are assigned to the most preferred natural science classes')
    # Determine the indices of the best matching natural science classes
    ns1_index = nat_sci_classes.index(best_ns_match_pair[0]) + 1
    ns2_index = nat_sci_classes.index(best_ns_match_pair[1]) + 1

    # Create a new class variable that will represent the class where students are assigned to the best matching natural science classes
    best_matching_class = model.NewIntVar(1, num_of_classes, 'best_matching_class')

    for student in students:
        # Create auxiliary boolean variables that are 1 if the student is assigned to the natural science classes ns1 and ns2 in any order
        is_assigned_ns1_ns2 = model.NewBoolVar(f'{student}_is_assigned_ns1_ns2')
        is_assigned_ns2_ns1 = model.NewBoolVar(f'{student}_is_assigned_ns2_ns1')

        # Ensure the natural science classes match the best match pair in any order
        model.Add(var_dict[str(student) + '_nat_sci_1'] == ns1_index).OnlyEnforceIf(is_assigned_ns1_ns2)
        model.Add(var_dict[str(student) + '_nat_sci_2'] == ns2_index).OnlyEnforceIf(is_assigned_ns1_ns2)
        model.Add(var_dict[str(student) + '_nat_sci_1'] == ns2_index).OnlyEnforceIf(is_assigned_ns2_ns1)
        model.Add(var_dict[str(student) + '_nat_sci_2'] == ns1_index).OnlyEnforceIf(is_assigned_ns2_ns1)

        # Ensure that if a student is in the best matching class, they must be assigned to the best match pair of natural science classes in any order
        is_in_best_matching_class = model.NewBoolVar(f'{student}_in_best_matching_class')
        model.Add(var_dict[str(student) + '_class'] == best_matching_class).OnlyEnforceIf(is_in_best_matching_class)
        model.Add(var_dict[str(student) + '_class'] != best_matching_class).OnlyEnforceIf(is_in_best_matching_class.Not())

        # This line ensures that if the student is in the best matching class, they must satisfy one of the conditions
        model.Add(is_assigned_ns1_ns2 + is_assigned_ns2_ns1 == is_in_best_matching_class)


    # cons_names.append('7. Students having language 5 must be in the same class')
    # first_student_with_lang_5_class = None
    # for student in students:
    #     has_lang_5 = model.NewBoolVar(f'{student}_has_lang_5')
    #     model.Add(var_dict[str(student) + '_lang'] == 5).OnlyEnforceIf(has_lang_5)
    #     model.Add(var_dict[str(student) + '_lang'] != 5).OnlyEnforceIf(has_lang_5.Not())
    #
    #     if first_student_with_lang_5_class is None:
    #         first_student_with_lang_5_class = var_dict[str(student) + '_class']
    #     else:
    #         model.Add(var_dict[str(student) + '_class'] == first_student_with_lang_5_class).OnlyEnforceIf(has_lang_5)

    # cons_names.append('7. Students having language 5 must be assigned to class 2')
    # for student in students:
    #     is_lang_5 = model.NewBoolVar(f'{student}_is_lang_5')
    #     is_class_2 = model.NewBoolVar(f'{student}_is_class_2')
    #
    #     model.Add(var_dict[str(student) + '_lang'] == 5).OnlyEnforceIf(is_lang_5)
    #     model.Add(var_dict[str(student) + '_lang'] != 5).OnlyEnforceIf(is_lang_5.Not())
    #
    #     model.Add(var_dict[str(student) + '_class'] == 2).OnlyEnforceIf(is_class_2)
    #     model.Add(var_dict[str(student) + '_class'] != 2).OnlyEnforceIf(is_class_2.Not())
    #
    #     model.Add(is_class_2 == is_lang_5)

    # cons_names.append('8. All male students must be in the same class')
    # male_students = data[data['Gender'] == 'm']['Student'].tolist()
    # if male_students:
    #     first_male_class = var_dict[str(male_students[0]) + '_class']
    #     for male_student in male_students[1:]:
    #         model.Add(var_dict[str(male_student) + '_class'] == first_male_class)

    cons_names.append(
        '8. Male students are split into exactly two classes, with each containing 4–7 males and one class containing none')

    male_students = data[data['Gender'] == 'm']['Student'].tolist()
    if male_students:

        # For each class, count how many male students are assigned there
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

        # Create boolean variables for classes that contain males
        classes_with_males = []
        for i in range(num_of_classes):
            has_males = model.NewBoolVar(f'class_{i + 1}_has_males')
            # has_males is true if male_count_in_class_i >= 1
            model.Add(male_counts_per_class[i] >= 1).OnlyEnforceIf(has_males)
            model.Add(male_counts_per_class[i] == 0).OnlyEnforceIf(has_males.Not())
            classes_with_males.append(has_males)

        # Enforce that exactly two classes have males
        model.Add(sum(classes_with_males) == 2)

        # Enforce that in those two classes, number of males is between 4 and 7
        for count in male_counts_per_class:
            # Either zero (no males) or between 4 and 7
            is_nonzero = model.NewBoolVar(f'{count.Name()}_nonzero')
            model.Add(count >= 1).OnlyEnforceIf(is_nonzero)
            model.Add(count == 0).OnlyEnforceIf(is_nonzero.Not())

            # Between 4 and 7 if nonzero
            model.Add(count >= 4).OnlyEnforceIf(is_nonzero)
            model.Add(count <= 7).OnlyEnforceIf(is_nonzero)

    cons_names.append('9. All student who chose language 4 as their first priority must be assigned to language 4')
    for student in students:
        highest_priority_lang = var_dict[str(student) + '_lang_priorities'][0]  # The highest priority language
        lang_4 = 4  # Language 4

        # If the highest priority language is language 4, then the chosen language must be language 4
        model.Add(var_dict[str(student) + '_lang'] == lang_4).OnlyEnforceIf(highest_priority_lang == lang_4)



    cons_names.append('10. There can be only max_class_size students assigned to the natural science class 2 (physics)')
    # Create a list to hold boolean variables indicating if a student is assigned to the specific natural science class
    students_assigned_to_ns2 = []

    # Define the index of the specific natural science class (e.g., 2)
    specific_ns_index = 2

    for student in students:
        # Create boolean variables for the conditions
        assigned_ns2_as_ns1 = model.NewBoolVar(f'{student}_assigned_ns2_as_ns1')
        assigned_ns2_as_ns2 = model.NewBoolVar(f'{student}_assigned_ns2_as_ns2')

        # Add constraints for these conditions
        model.Add(var_dict[str(student) + '_nat_sci_1'] == specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns1)
        model.Add(var_dict[str(student) + '_nat_sci_1'] != specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns1.Not())

        model.Add(var_dict[str(student) + '_nat_sci_2'] == specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns2)
        model.Add(var_dict[str(student) + '_nat_sci_2'] != specific_ns_index).OnlyEnforceIf(assigned_ns2_as_ns2.Not())

        # Collect these boolean variables
        students_assigned_to_ns2.append(assigned_ns2_as_ns1)
        students_assigned_to_ns2.append(assigned_ns2_as_ns2)

    # Add a constraint to limit the number of students assigned to the specific natural science class
    model.Add(sum(students_assigned_to_ns2) <= max_class_size)

    cons_names.append('11. All students with Physics as first or second priority must be assigned to Physics')
    physics_index = nat_sci_classes.index('Physics') + 1  # Convert to 1-based index (Physics == 2)
    for student in students:
        # Get the student's first and second priority natural science classes
        first_priority_ns = var_dict[str(student) + '_nat_sci_priorities'][0]  # Top priority
        second_priority_ns = var_dict[str(student) + '_nat_sci_priorities'][1]  # Second priority

        # Check if Physics is among top two priorities
        physics_top2 = model.NewBoolVar(f'{student}_physics_top2')
        model.AddBoolOr([
            first_priority_ns == physics_index,
            second_priority_ns == physics_index
        ]).OnlyEnforceIf(physics_top2)
        model.AddBoolAnd([
            first_priority_ns != physics_index,
            second_priority_ns != physics_index
        ]).OnlyEnforceIf(physics_top2.Not())

        # Force Physics to be assigned in one of the two slots if it's a top-2 priority
        has_physics_as_ns1 = model.NewBoolVar(f'{student}_assigned_physics_ns1')
        has_physics_as_ns2 = model.NewBoolVar(f'{student}_assigned_physics_ns2')

        model.Add(var_dict[str(student) + '_nat_sci_1'] == physics_index).OnlyEnforceIf(has_physics_as_ns1)
        model.Add(var_dict[str(student) + '_nat_sci_1'] != physics_index).OnlyEnforceIf(has_physics_as_ns1.Not())

        model.Add(var_dict[str(student) + '_nat_sci_2'] == physics_index).OnlyEnforceIf(has_physics_as_ns2)
        model.Add(var_dict[str(student) + '_nat_sci_2'] != physics_index).OnlyEnforceIf(has_physics_as_ns2.Not())

        # At least one of the two assignments must be Physics if top-2 priority
        model.AddBoolOr([has_physics_as_ns1, has_physics_as_ns2]).OnlyEnforceIf(physics_top2)

    cons_names.append('12. Students with Physics as 3rd priority must not be assigned to Physics')

    physics_index = nat_sci_classes.index('Physics') + 1  # Physics == 2

    for student in students:
        # Get student's 3rd priority natural science
        third_priority_ns = var_dict[str(student) + '_nat_sci_priorities'][2]  # 0-based indexing

        # Check if Physics is their 3rd priority
        physics_is_third_priority = model.NewBoolVar(f'{student}_physics_is_third_priority')
        model.Add(third_priority_ns == physics_index).OnlyEnforceIf(physics_is_third_priority)
        model.Add(third_priority_ns != physics_index).OnlyEnforceIf(physics_is_third_priority.Not())

        # If Physics is 3rd priority, enforce that neither slot is Physics
        model.Add(var_dict[str(student) + '_nat_sci_1'] != physics_index).OnlyEnforceIf(physics_is_third_priority)
        model.Add(var_dict[str(student) + '_nat_sci_2'] != physics_index).OnlyEnforceIf(physics_is_third_priority)

    cons_names.append('13. Any student assigned Physics must also be assigned Biology in the other slot')

    biology_index = nat_sci_classes.index('Biology') + 1  # Biology == 1
    physics_index = nat_sci_classes.index('Physics') + 1  # Physics == 2

    for student in students:
        ns1 = var_dict[str(student) + '_nat_sci_1']
        ns2 = var_dict[str(student) + '_nat_sci_2']

        # Create Boolean variables for "is Physics assigned to ns1 or ns2?"
        ns1_is_physics = model.NewBoolVar(f'{student}_ns1_is_physics')
        ns2_is_physics = model.NewBoolVar(f'{student}_ns2_is_physics')

        model.Add(ns1 == physics_index).OnlyEnforceIf(ns1_is_physics)
        model.Add(ns1 != physics_index).OnlyEnforceIf(ns1_is_physics.Not())

        model.Add(ns2 == physics_index).OnlyEnforceIf(ns2_is_physics)
        model.Add(ns2 != physics_index).OnlyEnforceIf(ns2_is_physics.Not())

        # If ns1 is Physics, enforce ns2 is Biology
        model.Add(ns2 == biology_index).OnlyEnforceIf(ns1_is_physics)

        # If ns2 is Physics, enforce ns1 is Biology
        model.Add(ns1 == biology_index).OnlyEnforceIf(ns2_is_physics)

    cons_names.append('14. Students with Spanish as top priority must not be assigned Italian')
    spanish_index = languages.index('Spanish') + 1  # Spanish == 5
    italian_index = languages.index('Italian') + 1  # Italian == 2
    for student in students:
        chosen_lang = var_dict[str(student) + '_lang']
        top_priority_lang = var_dict[str(student) + '_lang_priorities'][0]  # Top priority

        # Create a BoolVar: top priority is Spanish?
        top_priority_is_spanish = model.NewBoolVar(f'{student}_top_priority_is_spanish')
        model.Add(top_priority_lang == spanish_index).OnlyEnforceIf(top_priority_is_spanish)
        model.Add(top_priority_lang != spanish_index).OnlyEnforceIf(top_priority_is_spanish.Not())

        # Enforce: if top priority is Spanish, assigned language ≠ Italian
        model.Add(chosen_lang != italian_index).OnlyEnforceIf(top_priority_is_spanish)

    # ---------- Create the objective function -------------------

    # Objective function parameters
    lang_importance = 1
    lang_penalty = 10
    nat_sci_1_importance = 100
    nat_sci_2_importance = 1
    nat_sci_penalty = 100
    stratification = 2


    # Define the objective function
    objective_terms = []
    for student in students:

        chosen_lang = var_dict[str(student) + '_lang']
        chosen_nat_sci_1 = var_dict[str(student) + '_nat_sci_1']
        chosen_nat_sci_2 = var_dict[str(student) + '_nat_sci_2']


        for i in range(1, 6):
            lang_at_priority_index_i = var_dict[str(student) + '_lang_priorities'][i - 1]

            # Create a boolean variable that is 1 if chosen_lang == i
            assigned_lang_i = model.NewBoolVar(f'{student}_chosen_lang_{i}')
            model.Add(chosen_lang == lang_at_priority_index_i).OnlyEnforceIf(assigned_lang_i)
            model.Add(chosen_lang != lang_at_priority_index_i).OnlyEnforceIf(assigned_lang_i.Not())

            # Add the term to the objective function
            objective_terms.append(lang_importance * assigned_lang_i * (5 - i) ** stratification)

            # Penalize if the chosen language is not the one with the highest priority
            if i == 1:
                objective_terms.append(-lang_penalty * assigned_lang_i.Not())

            if i < 4:  # There are only 3 natural science classes

                # Create a boolean variable that is 1 if chosen_nat_sci_1 == i
                assigned_nat_sci_1 = model.NewBoolVar(f'{student}_chosen_nat_sci_1_{i}')
                model.Add(chosen_nat_sci_1 == lang_at_priority_index_i).OnlyEnforceIf(assigned_nat_sci_1)
                model.Add(chosen_nat_sci_1 != lang_at_priority_index_i).OnlyEnforceIf(assigned_nat_sci_1.Not())

                # Add the term to the objective function
                objective_terms.append(nat_sci_1_importance * assigned_nat_sci_1 * (3 - i) ** stratification)

                # Create a boolean variable that is 1 if chosen_nat_sci_2 == i
                assigned_nat_sci_2 = model.NewBoolVar(f'{student}_chosen_nat_sci_2_{i}')
                model.Add(chosen_nat_sci_2 == lang_at_priority_index_i).OnlyEnforceIf(assigned_nat_sci_2)
                model.Add(chosen_nat_sci_2 != lang_at_priority_index_i).OnlyEnforceIf(assigned_nat_sci_2.Not())

                # Add the term to the objective function
                objective_terms.append(nat_sci_2_importance * assigned_nat_sci_2 * (3 - i) ** stratification)

                # Penalize if the highest priority natural science classes are not chosen
                if i == 1:
                    objective_terms.append(-nat_sci_penalty * (assigned_nat_sci_1.Not() + assigned_nat_sci_2.Not()))


    # Maximize the sum of the objective terms
    model.Maximize(sum(objective_terms))

    # Create the solver and solve the model
    solver = cp_model.CpSolver()
    status = solver.Solve(model)

    # Check the result
    if status == cp_model.OPTIMAL or status == cp_model.FEASIBLE:
        print(f'\nTotal score: {solver.ObjectiveValue()}\n')

        if solver.ObjectiveValue() > best_score:
            best_score = solver.ObjectiveValue()
            best_data = data.copy()

            # Create new columns in the dataframe to store the results
            best_data['Class'] = None
            best_data['Language'] = None
            best_data['NatSci1'] = None
            best_data['NatSci2'] = None

            for student in students:

                # Add the results to the dataframe
                best_data.loc[data['Student'] == student, 'Class'] = solver.Value(var_dict[str(student) + "_class"])
                best_data.loc[data['Student'] == student, 'Language'] = languages[solver.Value(var_dict[str(student) + "_lang"]) - 1]
                best_data.loc[data['Student'] == student, 'NatSci1'] = nat_sci_classes[solver.Value(var_dict[str(student) + "_nat_sci_1"]) - 1]
                best_data.loc[data['Student'] == student, 'NatSci2'] = nat_sci_classes[solver.Value(var_dict[str(student) + "_nat_sci_2"]) - 1]
    else:
        print('No solution found.')

# Sort the data by class, then by language, then by natural science classes
best_data = best_data.sort_values(by=['Class', 'Language', 'NatSci1', 'NatSci2'])

# ---------------------------------------------------

# Class sizes
class_sizes = best_data['Class'].value_counts().sort_index()
print('\nClass sizes:')
print(class_sizes)

# Create a new folder for the results
results_name = input('Save into dir name (inside results dir): ')
if not os.path.exists('results/' + results_name):
    os.makedirs('results/' + results_name)

# Save the results to an Excel file
best_data.to_excel('results/' + results_name + '/' + filename + '_' + results_name + appendix, index=False)

# Save model parameters to a text file
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
    f.write(f'Best match pair of natural science classes: {best_ns_match_pair[0]} and {best_ns_match_pair[1]}')

