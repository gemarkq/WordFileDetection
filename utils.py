from enum import Enum
import numpy as np
from zhon.hanzi import punctuation

"""
return Type:
value: 1 represent word is equal to pattern
value: 2 represent word is partially equal to pattern (word is spelled wrong)
value: 3 represent word is different from pattern
"""
class WordStatus(Enum):
    WORD_SPELL_CORRECT = 1
    WORD_SPELL_WRONG = 2
    WORD_TOTALLY_DIFFERENT = 3

def CalEditDistance(word1, pattern):
    n = len(word1)
    m = len(pattern)
    dp = np.zeros((n + 1, m + 1))
    for i in range(n+1):
        dp[i][0] = i
    for j in range(m+1):
        dp[0][j] = j

    for i in range(1, n+1):
        for j in range(1, m+1):
            left = dp[i-1][j] + 1
            down = dp[i][j-1] + 1
            left_down = dp[i-1][j-1]
            if word1[i - 1] != pattern[j - 1]:
                left_down += 1
            dp[i][j] = min(left_down, left, down)
    return dp[n][m]

def IsThreeSame(word, pattern):
    n = len(word)
    m = len(pattern)
    if m >= 3 and n >= 3:
        return (word[0: 3] == pattern[0: -1] == pattern[-3: -1])
    

def IsWordValid(word, pattern):
    if (word == pattern):
        return WordStatus.WORD_SPELL_CORRECT
    
    is_three_same = IsThreeSame(word, pattern)
    edit_distance = CalEditDistance(word, pattern)

    if is_three_same and edit_distance <= 3:
        return WordStatus.WORD_SPELL_WRONG
    elif not is_three_same and edit_distance <= 2:
        return WordStatus.WORD_SPELL_WRONG
    else:
        return WordStatus.WORD_TOTALLY_DIFFERENT

def IsValidWordPunctuation(word):
    if word in punctuation:
        return False
    return True

def IsValidParaPunctuation(para):
    for word in para:
        if word in punctuation:
            return False
    return True

def IsValidWordPuctExclude(word):
    punct_exclude = "“”‘’–"
    for char in word:
        if char in punct_exclude:
            continue
        if char in punctuation:
            return False
    return True

def IsChineseWord(word):
    if '\u4e00' <= word <= '\u9fa5':
        return True
    return False

def IsConatinChinese(sentence):
    for word in sentence:
        if '\u4e00' <= word <= '\u9fa5':
            return True
    return False

def SeperateLine():
    print("\n---------------------------------------------------------------------------------------------\n")

