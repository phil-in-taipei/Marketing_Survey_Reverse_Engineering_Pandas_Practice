import numpy
import pandas
import csv
import openpyxl

pandas.options.mode.chained_assignment = None

# Input variables ---------------------------------------------------------------------------------------------------------------
df = pandas.read_csv('colgate.csv', low_memory=False)

brand_consideration_list = ['brandconsideration_Colgate', 'brandconsideration_eSensodyne', # boolean 1 or 0
                            'brandconsideration_Darlie', 'brandconsideration_Parodontax',
                            'brandconsideration_Oral-B', 'brandconsideration_others']

health_consideration_list = ['Whitening', 'Plaque', 'GumHealth', 'Desenstizing', 'BreathFresh', # boolean 1 or 0
                             'OtherhealthIssue', 'nocare']

viewed_ad_person_category_list = ['SeenBossAd', 'SeenColleagueAd', 'NoSeenAds'] # boolean 1 or 0

cross_ref_one_to_five_list = ['PreferenceEnhance', 'PurchaseIntention',  'ImpressionOfColgate', 'UnderstandEnhance'] #scale 1:5

last_purchase_category_list = ['Within1W', 'Within1M', 'Within3M', 'Within6M', 'NeverBuyWithin1Y', 'Morethan1Y', 'NeverBuy']

viewed_ad_media_category_list = ['ads_text', 'ads_picture  ', 'ads_video ', 'ads_Influencer', 'Video_instream'] #ad_category_list

age_category_list = ['13-19', '20-29', '30-39', '40-49', '50-59', '60-69']

gender_category_list = [0, 1] #male is 0, female is 1

# -------------------------------------------------------------------------------------------------------------------------------

# output variables --------------------------------------------------------------------------------------------------------------
brands = ['Colgate', 'Sensodyne', 'Darlie', 'Parodontax', 'Oral B', 'Other']
improvement_row_options = ['Will Improve', 'Might Improve', 'No Affect', 'Might Not Improve', 'Will Not Improve', 'Total']
seen_ad_options_media = ['ads_text', 'ads_picture  ', 'ads_video ', 'ads_Influencer', 'Video_instream', 'Total']
seen_ad_options_person = ['SeenBossAd', 'SeenColleagueAd', 'NoSeenAds', 'Total']
last_purchase_row_options = ['Within1W', 'Within1M', 'Within3M', 'Within6M', 'NeverBuyWithin1Y', 'Morethan1Y', 'NeverBuy', 'Total']
impression_row_options = ['Strong Impression', 'Impression', 'Unremarkable', 'No Impression', 'Absolutely No Impression', 'Total']
health_consideration_row_options = ['Whitening', 'Plaque', 'GumHealth', 'Desenstizing', 'BreathFresh', # boolean 1 or 0
                                    'OtherhealthIssue', 'nocare', 'total']
age_row_options = ['13-19', '20-29', '30-39', '40-49', '50-59', '60-69', 'total']
gender_row_options = ['Male', 'Female', 'Total']

sub_rows = ['Percent', 'Count']
brand_question = 'When selecting a toothpaste, which brands are your top choices?'

#--------------------------------------------------------------------------------------------------------------------------------

# Input functions ----------------------------------------------------------------------------------

def create_data_product_counts_1_5_scale(brand, category): # this works for scale of 1 to 5 questions
    product_counts = []                # returns descending list corresponding to value on scale
    for index in range(5, 0, -1):      # modify by setting the range as number of options
        brand_consideration_cross_ref = df[
            (df[brand] == 1) & (df[category] == index)] # change to df[category]
        product_counts.append(
            brand_consideration_cross_ref[brand].count()) # this adds to the count of items meeting the condition
    return product_counts # returns respective count of meeting condition for each brand



def create_data_product_counts_boolean_category(brand, category_list): # like multiple choice, but don't need 'category' argument
    product_counts = []                                                # pass in 'category' as ''
    for item in category_list:                                         # health concerns
        brand_consideration_cross_ref = df[                            # category of ad viewed -- person
            (df[brand] == 1) & (df[item] == 1)]                        # category of ad viewed -- media
        product_counts.append(
            brand_consideration_cross_ref[brand].count())
    return product_counts # returns respective count of meeting condition for each brand




def create_data_product_counts_multiple_choice_category(brand, category_list, category): # for 'Gender' (0 or 1)
    product_counts = []                                                                  # for Age
    for item in category_list:                                                           # for LastPurchaseTime
        brand_consideration_cross_ref = df[
            (df[brand] == 1) & (df[category] == item)]
        product_counts.append(
            brand_consideration_cross_ref[brand].count())
    return product_counts # returns respective count of meeting condition for each brand


def create_data_insert_list(product_counts):
    product_percentage_of_count = []
    insert_array = []
    for i in product_counts:
        percentage = (i / sum(product_counts)) * 100
        product_percentage_of_count.append(round(percentage,2))
        insert_array.append(round(percentage,2))
        insert_array.append(int(i))
    insert_array.append('Empty')
    insert_array.append('Empty')
    return insert_array


# output data and functions ---------------------------------------------------------------------------------------------------
# Main loop dictionaries
analysis_1 = {
    "sheet": "CT_0",
    "row_question": 'What is your preference for Colgate (after viewing ad)',
    "row_labels": improvement_row_options,
    "function_type": create_data_product_counts_1_5_scale,
    "row_variable_categories": [],
    "row_variable": cross_ref_one_to_five_list[0],
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_2 = {
    "sheet": "CT_1",
    "row_question": 'Willingness to buy or recommend others to buy increased?',
    "row_labels": improvement_row_options,
    "function_type": create_data_product_counts_1_5_scale,
    "row_variable_categories": [],
    "row_variable": cross_ref_one_to_five_list[1],
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_3 = {
    "sheet": "CT_2",
    "row_question": 'Which ad have you seen (boss, coworker, or none)',
    "row_labels": seen_ad_options_person,
    "function_type": create_data_product_counts_boolean_category,
    "row_variable_categories": viewed_ad_person_category_list,
    "row_variable": '',
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_4 = {
    "sheet": "CT_3",
    "row_question": 'What is your impression of Colgate?',
    "row_labels": impression_row_options,
    "function_type": create_data_product_counts_1_5_scale,
    "row_variable_categories": [],
    "row_variable": cross_ref_one_to_five_list[2],
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_5 = {
    "sheet": "CT_4",
    "row_question": 'What type of ad have you seen (media)?',
    "row_labels": seen_ad_options_media,
    "function_type": create_data_product_counts_boolean_category,
    "row_variable_categories": viewed_ad_media_category_list,
    "row_variable": '',
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_6 = {
    "sheet": "CT_5",
    "row_question": 'When did was the last time you purchased toothpaste?',
    "row_labels": last_purchase_row_options,
    "function_type": create_data_product_counts_multiple_choice_category,
    "row_variable_categories": last_purchase_category_list,
    "row_variable": 'LastPurchaseTime',
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_7 = {
    "sheet": "CT_6",
    "row_question": 'Which oral health concerns do you pay attention to?',
    "row_labels": health_consideration_row_options,
    "function_type": create_data_product_counts_boolean_category,
    "row_variable_categories": health_consideration_list,
    "row_variable": '',
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_8 = {
    "sheet": "CT_7",
    "row_question": 'What is your current age?',
    "row_labels": age_row_options,
    "function_type": create_data_product_counts_multiple_choice_category,
    "row_variable_categories": age_category_list,
    "row_variable": 'Age',
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_9 = {
    "sheet": "CT_8",
    "row_question": 'What is your gender?',
    "row_labels": gender_row_options,
    "function_type": create_data_product_counts_multiple_choice_category,
    "row_variable_categories": gender_category_list,
    "row_variable": 'Gender',
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}

analysis_10 = {
    "sheet": "CT_9",
    "row_question": 'What is your understanding of Colgate products?',
    "row_labels": improvement_row_options,
    "function_type": create_data_product_counts_1_5_scale,
    "row_variable_categories": [],
    "row_variable": cross_ref_one_to_five_list[3],
    "column_labels": brands,
    "column_variables": brand_consideration_list,
    "column_question":  brand_question
}
data_analysis = (analysis_1, analysis_2, analysis_3, analysis_4, analysis_5,
                 analysis_6, analysis_7, analysis_8, analysis_9, analysis_10)
# -------------------------------------------------------------------------------------------

def get_data_frame(options):
    outside = []
    for o in options:
        outside.append(o)
        outside.append(o)
    inside = []
    for index in range(len(options)):
        for i in sub_rows:
            inside.append(i)
    hierarchical_index = list(zip(outside, inside))

    hier_index = pandas.MultiIndex.from_tuples(hierarchical_index)
    return pandas.DataFrame('Empty', hier_index, brands)


def write_data_frame_sheet(data_frame, options, sheet,
                           create_product_counts_arg,
                           category_list=[], #blank list for 1 to 5
                           category=''): # blank string for boolean category
    # load values into data frame
    brand_index = 0
    rows_index_range = (len(options) - 1) * 2 # get number of original rows for loops (pre-totals)
    index_number_list = []

    for index in range(1, rows_index_range, 2):
        index_number_list.append(index)

    for brand in brand_consideration_list:

        if create_product_counts_arg == create_data_product_counts_1_5_scale:
            product_counts_list = create_product_counts_arg(brand=brand,
                                                            category=category)
        elif create_product_counts_arg == create_data_product_counts_boolean_category:
            product_counts_list = create_product_counts_arg(brand=brand,
                                                            category_list=category_list)
        else:
            product_counts_list = create_product_counts_arg(brand=brand,
                                                            category_list=category_list,
                                                            category=category) # can't pass in unneccessary  argument

        data_frame[brands[brand_index]] = create_data_insert_list(product_counts_list)
        brand_index += 1

    #calculate and write in row totals
    total_row_values = []
    for i in range(1, rows_index_range, 2):
        total_row_values.append(data_frame.iloc[i].sum())

    # calculate percent values and row totals
    complete_row_totals = []
    row_percent_values = []
    for value in total_row_values:
        percent_value = (value / sum(total_row_values)) * 100
        row_percent_values.append(round(percent_value, 2))
        complete_row_totals.append(round(percent_value, 2))
        complete_row_totals.append(value)
    complete_row_totals.append(100.00)
    complete_row_totals.append(sum(total_row_values))
    sum_of_total_row_values = sum(total_row_values)

    data_frame['Total'] = complete_row_totals

    # calculate column totals and insert
    for b in brands:
        new_data_frame = data_frame
        brand_column = new_data_frame[b]
        values = []
        for index in range(1, rows_index_range, 2): # change according to length of options -1
            values.append(brand_column[index])
        total_percent_row_index = rows_index_range
        total_values_row_index = rows_index_range + 1
        new_data_frame[b][total_values_row_index] = sum(values) # change according to length of options -1
        new_data_frame[b][total_percent_row_index] = round((sum(values) / sum_of_total_row_values) * 100, 2) # change according to length of options -1
        data_frame = new_data_frame

    data_frame.to_excel(writer, startrow=2, sheet_name=sheet)

# Main loop ------------------------------------------------------------------------------------------------------------------------

data_frame = pandas.DataFrame()
data_frame.to_excel('colgate_output.xlsx')
writer = pandas.ExcelWriter('colgate_output.xlsx', mode='a')
# for loop for all items in data_analysis list

for analysis_dict in data_analysis:
    analysis = analysis_dict

    data_frame = get_data_frame(analysis['row_labels']) # options will be in tuple that will be looped through

    if analysis['function_type'] == create_data_product_counts_1_5_scale:
        write_data_frame_sheet(data_frame, analysis['row_labels'],
                               create_product_counts_arg=analysis['function_type'],
                               category=analysis['row_variable'],
                               sheet=analysis['sheet'])
    elif analysis['function_type'] == create_data_product_counts_boolean_category:
        write_data_frame_sheet(data_frame, analysis['row_labels'],
                               create_product_counts_arg=analysis['function_type'],
                               category_list=analysis['row_variable_categories'],
                               sheet=analysis['sheet'])
    else:
        write_data_frame_sheet(data_frame, analysis['row_labels'],
                               create_product_counts_arg=analysis['function_type'],
                               category_list=analysis['row_variable_categories'],
                               category=analysis['row_variable'],
                               sheet=analysis['sheet'])

writer.save()

# another loop to write in titles for excel sheets

rb = openpyxl.load_workbook('colgate_output.xlsx')
for analysis_dict in data_analysis:
    sheet = rb[analysis_dict["sheet"]] # sheet will be created according to CT_ 'index' in loop
    sheet['A1'] = str(analysis_dict["row_question"] + '/' + analysis_dict['column_question'])
    #sheet = rb["CT_6"] # sheet will be created according to CT_ 'index' in loop
    #sheet['A1'] = str(analysis['row_question'] + '/' + analysis['column_question'])
rb.save('colgate_output.xlsx')

# end main loop ---------------------------------------------------------------
