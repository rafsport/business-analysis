##### Per togliere le duplicazioni dalle stringhe #####

from collections import Counter

def remov_duplicates(input):

    # split input string separated by space
    input = input.split(" + ")

    # joins two adjacent elements in iterable way
    '''for i in range(0, len(input)):
        input[i] = "".join(input[i])
'''
    # now create dictionary using counter method
    # which will have strings as key and their
    # frequencies as value
    UniqW = Counter(input)

    # joins two adjacent elements in iterable way
    s = " + ".join(UniqW.keys())
    return s


##### Per esprimere il formato dei numeri #####
def human_format(num):
    num = float('{:.3g}'.format(num))
    magnitude = 0
    while abs(num) >= 1000:
        magnitude += 1
        num /= 1000.0
    return '{}{}'.format('{:f}'.format(num).rstrip('0').rstrip('.'), ['', 'k', 'M', 'B', 'T'][magnitude])


##### Highlights max value in a DataFrame #####
def highlight_max(x, color):
    return np.where(x == np.nanmax(x.to_numpy()), f"background: {color};", None)


##### Funzione per rinominare le colonne mantenendo solo la parte tra virgolette #####
def rename_columns(df):
    new_column_names = []
    for col in df.columns:
        # Trova la prima occorrenza di testo tra virgolette
        start = col.find('"')
        end = col.find('"', start + 1)
        if start != -1 and end != -1:
            new_column_names.append(col[start+1:end])
        else:
            new_column_names.append(col)
    df.columns = new_column_names