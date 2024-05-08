# pip.main(['install', 'package_name'])
import pandas as pd
from collections import OrderedDict
from matplotlib import pyplot as plt


# compare voice lines and identify duplicates
def compare_lines(list_lines_a, list_lines_b):
    duplicates = []
    for line_a in list_lines_a:
        if line_a in list_lines_b:
            duplicates.append(line_a)
    return duplicates


def ex_compare(list_a, list_b):
    ex_dupl = []
    for k in range(0, len(list_a)):
        if len(ex_dupl) != 0:
            if list_b.index(list_a[k-1]) < list_b.index(list_a[k]):
                ex_dupl.append(list_a[k])
        else:
            ex_dupl.append(list_a[k])
    return ex_dupl


# cluster voice lines into interactions
def cluster_interaction(list_lines, dataframe):
    interactions = []
    # counter for how many voice lines per interaction
    counter = []
    df = pd.DataFrame(columns=["Sprecher:in", "Angesprochene:r 1", "Angesprochene:r 2", "Interaktionstyp",
                               "Ort", "Text"])
    for line in list_lines:
        # find voice line in dataframe
        result = dataframe[dataframe["Text"] == line]
        x = dataframe.index.get_loc(result.index[0])

        df.loc[len(df.index)] = [dataframe["Sprecher:in"].iloc[x], dataframe["Angesprochene:r 1"].iloc[x],
                                 dataframe["Angesprochene:r2"].iloc[x], dataframe["Interaktionstyp"].iloc[x],
                                 dataframe["Ort"].iloc[x], dataframe["Text"].iloc[x]]

        # form lists from who speaks to whom, in which context and where -> establish one interaction
        speaker = str(result.at[x, "Sprecher:in"])
        addressed1 = str(result.at[x, "Angesprochene:r 1"])
        # typo in column name in source file, too lazy to change
        addressed2 = str(result.at[x, "Angesprochene:r2"])
        interaction = str(result.at[x, "Interaktionstyp"])
        place = str(result.at[x, "Ort"])
        t_list = [speaker, addressed1, addressed2, interaction, place]

        # first interaction has no check and just gets appended
        if len(interactions) == 0:
            interactions.append(t_list)
            counter.append(1)
        else:
            # check if next voice line occurs during the same interaction, if not append to interactions list
            if set(t_list) != set(interactions[len(interactions)-1]):
                interactions.append(t_list)
                counter.append(1)
            else:
                counter[len(counter)-1] = counter[len(counter)-1] + 1

    # merge counter with corresponding interaction
    for y in range(0, len(counter)):
        interactions[y].append(counter[y])

    return interactions, df


def plot_style(x_value, y_values, size_x, size_y, label_x, label_y, rot_degree):
    fig, ax = plt.subplots(figsize=(size_x, size_y))
    bars = ax.bar(x_value, y_values, color="gray", width=0.4)
    for bar in bars:
        height = bar.get_height()
        ax.annotate(f'{height}', xy=(bar.get_x() + bar.get_width() / 2, height), xytext=(0, 3),
                    textcoords="offset points", ha='center', va='bottom')
    ax.grid(color='grey', linestyle='-.', linewidth=0.5, alpha=0.2)
    plt.xticks(rotation=rot_degree)
    plt.xlabel(label_x)
    plt.ylabel(label_y)
    plt.show()


def series_to_values(input_series):
    x_values = input_series.index.tolist()
    y_values = input_series.tolist()
    return x_values, y_values


if __name__ == '__main__':
    # PREPARING DATA
    # read ods document into pandas dataframes
    df_ak = pd.read_excel("akuyaku_dialog.ods", engine="odf")
    df_dl = pd.read_excel("daelagor_dialog.ods", engine="odf")
    df_gg = pd.read_excel("graumengaming_dialog.ods", engine="odf")

    # turn voice lines into list
    text_ak_list = df_ak["Text"].tolist()
    text_dl_list = df_dl["Text"].tolist()
    text_gg_list = df_gg["Text"].tolist()

    # compare ak to dl
    dupl_ak_dl = compare_lines(text_ak_list, text_dl_list)
    # compare gg to ak and dl
    dupl_ak_dl_gg = compare_lines(dupl_ak_dl, text_gg_list)
    # remove duplicates in list
    dupl_ak_dl_gg = list(OrderedDict.fromkeys(dupl_ak_dl_gg))

    # INTERACTIONS IN EXACT SAME ORDER
    # compare both directions ak vs dl
    dupl_dl_ak = compare_lines(text_dl_list, text_ak_list)
    # remove duplicates from list
    dupl_ak_dl = list(OrderedDict.fromkeys(dupl_ak_dl))
    dupl_dl_ak = list(OrderedDict.fromkeys(dupl_dl_ak))
    # use method to check order
    ex_dl_ak = ex_compare(dupl_ak_dl, dupl_dl_ak)
    # remove duplicates
    text_gg_list = list(OrderedDict.fromkeys(text_gg_list))
    # compare both directions ak + dl vs. gg
    dupl_gg_ak_dl = compare_lines(text_gg_list, ex_dl_ak)
    dupl_dl_ak_gg = compare_lines(ex_dl_ak, text_gg_list)
    # use method to check order
    ex_dupl_gg_ak_dl = ex_compare(dupl_dl_ak_gg, dupl_gg_ak_dl)

    # cluster voice lines into interactions into dataframe
    list_interactions, df_dupl_lines = cluster_interaction(dupl_ak_dl_gg, df_ak)
    df_interactions = pd.DataFrame(list_interactions)
    df_interactions = df_interactions.set_axis(
        ["speaker", "addressed1", "addressed2", "type", "place", "no_lines"],
        axis="columns")

    # cluster EXACT SAME (in order) voice lines into interactions
    list_ex_interactions, df_ex_dupl_lines = cluster_interaction(ex_dupl_gg_ak_dl, df_ak)
    df_ex_interactions = pd.DataFrame(list_ex_interactions)
    df_ex_interactions = df_ex_interactions.set_axis(
        ["speaker", "addressed1", "addressed2", "type", "place", "no_lines"],
        axis="columns")

    # VOICE LINES PER INTERACTION
    # count how often an interaction contains n number of voice lines
    count_no_of_lines = df_interactions["no_lines"].value_counts()
    # also for EXACT SAME voice lines
    count_ex_no_of_lines = df_ex_interactions["no_lines"].value_counts()
    # gather complete information
    keys_lines = []
    values_lines = []
    for i in range(1, 18):
        keys_lines.append(str(i))
        if i in count_no_of_lines.index:
            values_lines.append(count_no_of_lines[i])
        else:
            values_lines.append(0)
    # complete information for EXACT SAME voice lines
    keys_ex_lines = []
    values_ex_lines = []
    for i in range(1, 18):
        keys_ex_lines.append(str(i))
        if i in count_ex_no_of_lines.index:
            values_ex_lines.append(count_ex_no_of_lines[i])
        else:
            values_ex_lines.append(0)

    # WHO SPEAKS
    count_speaker = df_dupl_lines["Sprecher:in"].value_counts()
    speaker_x, speaker_y = series_to_values(count_speaker)
    # also for exact same voice lines
    count_ex_speaker = df_ex_dupl_lines["Sprecher:in"].value_counts()
    ex_speaker_x, ex_speaker_y = series_to_values(count_ex_speaker)

    # TO WHOM
    count_addressed = df_dupl_lines[["Angesprochene:r 1", "Angesprochene:r 2"]].stack().value_counts()
    addressed_x, addressed_y = series_to_values(count_addressed)
    # also for exact same voice lines
    count_ex_addressed = df_ex_dupl_lines[["Angesprochene:r 1", "Angesprochene:r 2"]].stack().value_counts()
    ex_addressed_x, ex_addressed_y = series_to_values(count_ex_addressed)

    # TYPE OF INTERACTION
    count_type = df_interactions["type"].value_counts()
    type_x, type_y = series_to_values(count_type)
    # also for exact same voice lines
    count_ex_type = df_ex_interactions["type"].value_counts()
    ex_type_x, ex_type_y = series_to_values(count_ex_type)

    # PLACE
    count_place = df_interactions["place"].value_counts()
    place_x, place_y = series_to_values(count_place)
    # also for exact same voice lines
    count_ex_place = df_ex_interactions["place"].value_counts()
    ex_place_x, ex_place_y = series_to_values(count_ex_place)

    # CROSSTABLES
    speaker_x_type = pd.crosstab(df_interactions["speaker"], df_interactions["type"],
                                 margins=True, margins_name="Gesamt")
    speaker_x_place = pd.crosstab(df_interactions["speaker"], df_interactions["place"],
                                  margins=True, margins_name="Gesamt")
    interaction_x_place = pd.crosstab(df_interactions["type"], df_interactions["place"],
                                      margins=True, margins_name="Gesamt")
    # also for exact same voice lines
    EX_speaker_x_type = pd.crosstab(df_ex_interactions["speaker"], df_ex_interactions["type"],
                                    margins=True, margins_name="Gesamt")
    EX_speaker_x_place = pd.crosstab(df_ex_interactions["speaker"], df_ex_interactions["place"],
                                     margins=True, margins_name="Gesamt")
    EX_interaction_x_place = pd.crosstab(df_ex_interactions["type"], df_ex_interactions["place"],
                                         margins=True, margins_name="Gesamt")

    # checks for specific interactions
    """desired_value_index = 5
    length_of_interaction = 7
    for list_x in list_interactions:
        if list_x[desired_value_index] == length_of_interaction:
            print(list_x)"""

    # BARPLOTS
    plot_style(place_x, place_y, 10, 5,
               "Ort der Interaktion", "Häufigkeit", 90)
    plot_style(type_x, type_y, 10, 7,
               "Art der Interaktion", "Häufigkeit", 90)
    plot_style(addressed_x, addressed_y, 10, 7,
               "Zuhörende Figur", "Anzahl gehörter Sprechakte", 90)
    plot_style(speaker_x, speaker_y, 10, 7,
               "Sprechende Figur", "Anzahl Sprechakte", 90)
    plot_style(keys_lines, values_lines, 10, 5,
               "Sprechakte pro Interaktion", "Häufigkeit", 0)
    # also for exact same voice lines
    plot_style(ex_place_x, ex_place_y, 10, 5,
               "Ort der Interaktion", "Häufigkeit", 90)
    plot_style(ex_type_x, ex_type_y, 10, 7,
               "Art der Interaktion", "Häufigkeit", 90)
    plot_style(ex_addressed_x, ex_addressed_y, 10, 7,
               "Zuhörende Figur", "Anzahl gehörter Sprechakte", 90)
    plot_style(ex_speaker_x, ex_speaker_y, 10, 7,
               "Sprechende Figur", "Anzahl Sprechakte", 90)
    plot_style(keys_ex_lines, values_ex_lines, 10, 5,
               "Sprechakte pro Interaktion", "Häufigkeit", 0)

    # write interaction lists into text file
    with open(r'ex_interactions.txt', 'w') as f_ex:
        for m in list_ex_interactions:
            f_ex.write("%s\n" % m)
    with open(r'interactions.txt', 'w') as f:
        for n in list_interactions:
            f.write("%s\n" % n)

    print("Das Hades Let's Play von Akuyaku enthält", len(text_ak_list), "Voicelines.")
    print("Das Hades Let's Play von Daelagor enthält", len(text_dl_list), "Voicelines.")
    print("Das Hades Let's Play von GraumenGaming enthält", len(text_gg_list), "Voicelines.")
    print(len(dupl_ak_dl_gg), "der Voicelines kommen in allen drei Let's Plays vor.")
    print("Diese Voicelines können in", len(list_interactions), "Interaktionen gegliedert werden.")
    print(len(ex_dupl_gg_ak_dl), "der Voicelines kommen in der exakt gleichen Reihenfolge vor und können in",
          len(list_ex_interactions), "Interaktionen gegliedert werden.")
