# -*- coding: utf-8 -*-
import sqlite3


def opening_the_database():
    """Открытие базы данных"""
    conn = sqlite3.connect('data.db')  # Создаем соединение с базой данных
    cursor = conn.cursor()
    return conn, cursor


if __name__ == '__main__':
    opening_the_database()
