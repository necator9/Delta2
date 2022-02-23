import re
import glob
import os


def get_modules_list(old_paths_list, re_exp):
	filenames = list(os.path.basename(path) for path in old_paths_list)

	filtered_names = list()

	for line in filenames:
		matchObj = re.match(re_exp, line, re.M|re.I)
		extracted = matchObj.group()
		if matchObj and not extracted.endswith('.xlsx') :
		   filtered_names.append(extracted)

	return filtered_names



def find_matching_index(list1, list2):
    inverse_index = { element: index for index, element in enumerate(list1) }

    return [(index, inverse_index[element])
        for index, element in enumerate(list2) if element in inverse_index]

old_dir = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R4_4_PH_CodeBeamer\*.xlsx'
new_dir = r'C:\Users\LT45641\Documents\HVLM\Test\E3C_R5.2\Ivan\PH\R5_2_PH_CodeBeamer\*.xlsx'
re_exp = r'([^\s]+)'

old_paths_list = glob.glob(old_dir)
new_paths_list = glob.glob(new_dir)

old_names = get_modules_list(old_paths_list, re_exp)
f


new_names = get_modules_list(new_paths_list, re_exp)
print(old_names)
print(new_names)

print(find_matching_index(old_names, new_names))