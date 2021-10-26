import os
import time

import numpy as np
import pandas as pd
from openpyxl import Workbook, load_workbook


def new_xlsx(xlsx_title):
    date = time.strftime('%Y-%m-%d_%H-%M', time.localtime())
    xlsx_file_fame = "t4_readable_file\\" + xlsx_title + date + ".xlsx"
    # Create a xlsx Workbook
    wb = Workbook()
    ws = wb.active
    ws.title = xlsx_title
    wb.save(xlsx_file_fame)
    # open the xlsx
    load_workbook(xlsx_file_fame)
    return xlsx_file_fame


def new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions, text_Text,
                text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight, text_Assessment_team_info,
                text_Response_team_info, df):
    new_row = pd.DataFrame({'Questionnaire': [text_Questionnaire],
                            'Description': [text_Description],
                            'Criteria': [text_Criteria],
                            'Parent Criterion': [text_Parent_Criterion],
                            'Questions': [text_Questions],
                            'Text': [text_Text],
                            'Legends': [text_Legends],
                            'Possible Answers': [text_Possible_Answers],
                            'P.A Legend': [text_PA_Legend],
                            'Type': [text_Type],
                            'Weight': [text_Weight],
                            'Assessment team info': [text_Assessment_team_info],
                            'Response team info': [text_Response_team_info]})

    df = pd.concat([new_row, df]).reset_index(drop=True)
    return df


def refactor_file(df, answer_type):
    # delete unnecessary columns
    df.drop('Unnamed: 7', axis=1, inplace=True)
    df.drop('n: notwendig', axis=1, inplace=True)
    # delete unnecessary rows
    df.drop(index=0, axis=0, inplace=True)

    # rename columns
    df = df.rename(columns={'Prüfpunkte Titel (n)': "Questions"})
    df = df.rename(columns={'Prüfpunkt Frage/Anforderung (n)': "Text"})
    df = df.rename(columns={'Unterkriterium (o)': "Criteria"})
    df = df.rename(columns={'Kriterium (n)': "Parent Criterion"})
    df = df.rename(columns={'Information Auditor, z.B. Anforderung aus NORM (o)': "Assessment team info"})
    df = df.rename(columns={'Information Auditee (o), z.B. Ziel': "Response team info"})
    # df = df.rename(columns={'': ""})

    # sort columns in correct order
    df = df.reindex(columns=['Questionnaire',
                             'Description',
                             'Criteria',
                             'Parent Criterion',
                             'Questions',
                             'Text',
                             'Legends',
                             'Possible Answers',
                             'P.A Legend',
                             'Type',
                             'Weight',
                             'Assessment team info',
                             'Response team info',
                             'Norm-Ref. (o)'])

    # set missing parameters

    for i in range(1, len(df) + 1):
        # set Parent Criterion as Criteria if Criteria not exist
        # if pd.isnull(df.at[i, 'Criteria']) is True:
        #     df.loc[[i], 'Criteria'] = df.at[i, 'Parent Criterion']
        # set default Weight
        if pd.isnull(df.at[i, 'Questions']) is False:
            if pd.isnull(df.at[i, 'Weight']) is True:
                df.loc[[i], 'Weight'] = 1
        # set ANSWER_TYPE
        if pd.isnull(df.at[i, 'Questions']) is False:
            if pd.isnull(df.at[i, 'Possible Answers']) is True:
                df.loc[[i], 'Possible Answers'] = "ANSWER_TYPE_" + str(answer_type)

    # #df.loc[[i], 'Possible Answers'] = "ANSWER_TYPE_" + str(answerType)

    for i in reversed(range(1, len(df) + 1)):
        if pd.isnull(df.at[i, 'Parent Criterion']) is False:

            text_Questionnaire = ""
            text_Description = ""
            text_Criteria = df.at[i, 'Parent Criterion']
            text_Parent_Criterion = ""
            text_Questions = ""
            text_Text = ""
            text_Legends = ""
            text_Possible_Answers = ""
            text_PA_Legend = ""
            text_Type = ""
            text_Weight = ""
            text_Assessment_team_info = ""
            text_Response_team_info = ""
            df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                             text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                             text_Assessment_team_info, text_Response_team_info, df)
            new_row_order = []
            # for j in range(1, i):
            #     new_row_order.append(j)
            #     print(j)
            # new_row_order.append(0)
            # for z in range(i + 1, len(df) + 1):
            #     new_row_order.append(z)
            # df = df.reindex(index=new_row_order)

    return df


def define_answer_rows(answer_type, df):
    if answer_type == 1:
        text_Questionnaire = ""
        text_Description = ""
        text_Criteria = ""
        text_Parent_Criterion = ""
        text_Questions = ""
        text_Text = ""
        text_Legends = ""
        text_Possible_Answers = "n.a."
        text_PA_Legend = "nicht anwendbar"
        text_Type = "IMP"
        text_Weight = ""
        text_Assessment_team_info = "nicht anwendbar"
        text_Response_team_info = ""
        df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                         text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                         text_Assessment_team_info, text_Response_team_info, df)

        text_Questionnaire = ""
        text_Description = ""
        text_Criteria = ""
        text_Parent_Criterion = ""
        text_Questions = ""
        text_Text = ""
        text_Legends = ""
        text_Possible_Answers = "nicht erfüllt"
        text_PA_Legend = "0% - 15%"
        text_Type = "NOR"
        text_Weight = 100
        text_Assessment_team_info = "Es gibt wenig bis keinen Nachweis, dass ein Attribut eines Prozesses erfüllt wird. (0% - 15%)"
        text_Response_team_info = "Es gibt wenig bis keinen Nachweis, dass ein Attribut eines Prozesses erfüllt wird. (0% - 15%)"
        df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                         text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                         text_Assessment_team_info, text_Response_team_info, df)

        text_Questionnaire = ""
        text_Description = ""
        text_Criteria = ""
        text_Parent_Criterion = ""
        text_Questions = ""
        text_Text = ""
        text_Legends = ""
        text_Possible_Answers = "teilweise erfüllt"
        text_PA_Legend = "16% - 50%"
        text_Type = "NOR"
        text_Weight = 67
        text_Assessment_team_info = "Es gibt einen Nachweis, dass ein Attribut eines Prozesses teilweise erfüllt wird. (15% - 50%)"
        text_Response_team_info = "Es gibt einen Nachweis, dass ein Attribut eines Prozesses teilweise erfüllt wird. (15% - 50%)"
        df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                         text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                         text_Assessment_team_info, text_Response_team_info, df)

        text_Questionnaire = ""
        text_Description = ""
        text_Criteria = ""
        text_Parent_Criterion = ""
        text_Questions = ""
        text_Text = ""
        text_Legends = ""
        text_Possible_Answers = "weitgehend erfüllt"
        text_PA_Legend = "51% - 85%"
        text_Type = "NOR"
        text_Weight = 32
        text_Assessment_team_info = "Es gibt einen Nachweis eines systematischen Ansatzes und des signifikanten Erfüllens eines Attributs eines Prozesses. Es können Schwächen bzgl. des Attributs vorliegen. (50% - 85%"
        text_Response_team_info = "Es gibt einen Nachweis eines systematischen Ansatzes und des signifikanten Erfüllens eines Attributs eines Prozesses. Es können Schwächen bzgl. des Attributs vorliegen. (50% - 85%"
        df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                         text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                         text_Assessment_team_info, text_Response_team_info, df)

        text_Questionnaire = ""
        text_Description = ""
        text_Criteria = ""
        text_Parent_Criterion = ""
        text_Questions = ""
        text_Text = ""
        text_Legends = ""
        text_Possible_Answers = "vollständig erfüllt"
        text_PA_Legend = "86% - 100%"
        text_Type = "NOR"
        text_Weight = 0
        text_Assessment_team_info = "Es gibt einen Nachweis eines vollständigen, systematischen Ansatzes und vollständig erfüllten Attributs eines Prozesses. Keine nennenswerten Schwächen bzgl. des Attributs liegen vor. (85% - 100%)"
        text_Response_team_info = "Es gibt einen Nachweis eines vollständigen, systematischen Ansatzes und vollständig erfüllten Attributs eines Prozesses. Keine nennenswerten Schwächen bzgl. des Attributs liegen vor. (85% - 100%)"
        new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                    text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                    text_Assessment_team_info, text_Response_team_info, df)

    # order rows
    new_row_order = [4, 0, 1, 2, 3, 5]
    for i in range(5, len(df) + 1):
        new_row_order.append(i)
    df = df.reindex(index=new_row_order)

    return df


# Add 3 rows on top
def define_damage_and_name_rows(df):
    text_Questionnaire = "Mittel"
    text_Description = 25000
    text_Criteria = ""
    text_Parent_Criterion = ""
    text_Questions = ""
    text_Text = ""
    text_Legends = ""
    text_Possible_Answers = ""
    text_PA_Legend = ""
    text_Type = ""
    text_Weight = ""
    text_Assessment_team_info = ""
    text_Response_team_info = ""
    df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                     text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                     text_Assessment_team_info, text_Response_team_info, df)

    text_Questionnaire = "Hoch"
    text_Description = 50000
    text_Criteria = ""
    text_Parent_Criterion = ""
    text_Questions = ""
    text_Text = ""
    text_Legends = ""
    text_Possible_Answers = ""
    text_PA_Legend = ""
    text_Type = ""
    text_Weight = ""
    text_Assessment_team_info = ""
    text_Response_team_info = ""
    df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                     text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                     text_Assessment_team_info, text_Response_team_info, df)

    text_Questionnaire = "Gering"
    text_Description = 5000
    text_Criteria = ""
    text_Parent_Criterion = ""
    text_Questions = ""
    text_Text = ""
    text_Legends = ""
    text_Possible_Answers = ""
    text_PA_Legend = ""
    text_Type = ""
    text_Weight = ""
    text_Assessment_team_info = ""
    text_Response_team_info = ""
    df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                     text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                     text_Assessment_team_info, text_Response_team_info, df)

    text_Questionnaire = "!AUSFÜLLEN"
    text_Description = "Beschreibung Vorlage Assessment"
    text_Criteria = ""
    text_Parent_Criterion = ""
    text_Questions = ""
    text_Text = ""
    text_Legends = ""
    text_Possible_Answers = ""
    text_PA_Legend = ""
    text_Type = ""
    text_Weight = ""
    text_Assessment_team_info = "# Anleitung und Erklärungen zum vorliegenden Assessment"
    text_Response_team_info = "# Anleitung und Erklärungen zum vorliegenden Assessment"
    df = new_row_top(text_Questionnaire, text_Description, text_Criteria, text_Parent_Criterion, text_Questions,
                     text_Text, text_Legends, text_Possible_Answers, text_PA_Legend, text_Type, text_Weight,
                     text_Assessment_team_info, text_Response_team_info, df)
    return df


def generate_file(df, xlsx_file_fame):
    print(df)
    cols = list(df.columns.values)
    print(cols)
    df.to_excel(xlsx_file_fame, index=False, header=True)


def convert_to_t4_excel(xlsx_title, answer_type, inputFilename):
    # open file
    df: object = pd.read_excel(inputFilename)

    xlsx_file_fame = new_xlsx(xlsx_title)

    df = refactor_file(df, answer_type)

    # df = define_answer_rows(answer_type, df)

    # df = define_damage_and_name_rows(df)

    generate_file(df, xlsx_file_fame)


def main():
    xlsx_title = "main"
    answer_type = 1
    inputFilename = "C:\\Users\\Mika F\\Documents\\etoe\\Template_Fragebogen_K4_Design.xlsx"

    convert_to_t4_excel(xlsx_title, answer_type, inputFilename)


if __name__ == '__main__':
    main()
