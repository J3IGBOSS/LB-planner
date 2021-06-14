import pickle


def save_pkl(data, filename):
    with open(filename, 'wb') as f:
        pickle.dump(data, f)
    # print(f'Complete save of {filename}.')


def retrieve_pkl(filename):
    with open(filename, 'rb') as f:
        data = pickle.load(f)
    return data
