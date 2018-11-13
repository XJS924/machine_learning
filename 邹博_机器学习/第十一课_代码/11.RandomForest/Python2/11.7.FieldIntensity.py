#coding:utf-8

import numpy as np
import pandas as pd
from sklearn.ensemble import RandomForestRegressor
from sklearn.model_selection import train_test_split
from sklearn.metrics import r2_score
from sklearn.model_selection import GridSearchCV
import matplotlib.pyplot as plt


def show_data(y, y_pred, title):
    plt.figure(figsize=(7, 6), facecolor='w')
    plt.plot(y, 'r-', lw=2, label='Actual')
    plt.plot(y_pred, 'g-', lw=1, label='Predict', alpha=0.7)
    plt.grid(True)
    plt.xlabel('Samples', fontsize=15)
    plt.ylabel('Field Intensity', fontsize=15)
    plt.legend(loc='upper left')
    plt.title(title, fontsize=18)
    plt.tight_layout()


if __name__ == '__main__':
    data_prime = pd.read_csv('FieldIntensity.csv', header=0)
    data_group = data_prime.groupby(by=['x', 'y'])
    data_mean = data_group.mean()
    data = data_mean.reset_index()
    print data

    x = data[['x', 'y']]
    y = data['88000KHz']
    x_train, x_test, y_train, y_test = train_test_split(x, y, train_size=0.75, random_state=0)
    # rf = RandomForestRegressor(n_estimators=100, criterion='mse', max_depth=11, min_samples_split=3)
    # model = GridSearchCV(rf, param_grid={'max_depth': np.arange(10, 15), 'min_samples_split': np.arange(1, 5)})
    model = RandomForestRegressor(n_estimators=100, criterion='mse', max_depth=14, min_samples_split=3)
    model.fit(x_train, y_train)

    order = y_train.argsort(axis=0)
    y_train = y_train.values[order]
    x_train = x_train.values[order, :]
    y_train_pred = model.predict(x_train)

    order = y_test.argsort(axis=0)
    y_test = y_test.values[order]
    x_test = x_test.values[order, :]
    y_test_pred = model.predict(x_test)

    print r2_score(y_train, y_train_pred)
    print r2_score(y_test, y_test_pred)

    show_data(y_train, y_train_pred, 'Train Data')
    show_data(y_test, y_test_pred, 'Test Data')
    plt.show()
