import pytest

from secret_santa import RandomDraw

import numpy as np
from collections import Counter
from itertools import permutations
from math import factorial

@pytest.fixture
def initStdDraw():
    """
    Initialize a RandomDraw object with a specific occurrence matrix for testing.
    """
    draw = RandomDraw()
    draw.m_occurrence = np.array([["Dipsy", "dispy@mail.com", 0, 1, 0, 2], 
                                 ["Laa-Laa", "laalaa@mail.com", 0, 0, 2, 3], 
                                 ["Po", "po@mail.com", 3, 0, 0, 1], 
                                 ["Tinky-Winky", "tinky.winky@mail.com", 0, 2, 2, 0]])
    draw.m_avoidance = np.array([])
    draw.m_names = list(draw.m_occurrence[:,0])
    draw.m_mails = dict(zip(draw.m_names, list(draw.m_occurrence[:,1])))
    draw.resetNameAvailability()
    return draw

def test_drawAll(initStdDraw):
    """
    Test that the draw is valid : same number of names and mails, every one is assigned, 
    and no one draws themselves, 
    """
    draw = initStdDraw

    for _ in range(500):
        result = draw.process()
        
        assert len(result) == len(draw.m_names), "The result is the same length as the number of names."
        assert len(result) == len(draw.m_mails), "The result is the same length as the number of mails."
        assert all(name in result for name in draw.m_names), "Every name in draw.m_names is present in the results values."
        assert all(result[name] != name for name in draw.m_names), "No one draws themselves."

def test_possibilityList(initStdDraw):
    """
    Test that the possibility list is valid : no one has himself in the list, all names in the list are available, and the number of occurrences is correct.
    """
    draw = initStdDraw

    # Make possibility list when every name is available
    for testIndex, _ in enumerate(draw.m_names):
        name = draw.m_names[testIndex]
        possibilities = draw.makePossibilityList(testIndex)

        assert name not in possibilities, "The person should not be able to draw themselves."
        assert all(draw.m_nameAvailable[name] for name in possibilities), "All names in the possibility list should be available."

        occurrences = draw.m_occurrence[testIndex, 2:].astype(int)
        maxOcc = np.max(occurrences)+1
        for nameIndex, _ in enumerate(draw.m_names):
            if nameIndex != testIndex:
                assert possibilities.count(draw.m_names[nameIndex]) == (maxOcc - occurrences[nameIndex])

    # Possibility list when one is not available
    draw.m_nameAvailable[draw.m_names[1]] = False
    for testIndex, _ in enumerate(draw.m_names):
            name = draw.m_names[testIndex]
            possibilities = draw.makePossibilityList(testIndex)
    
            assert name not in possibilities, "The person should not be able to draw themselves."
            assert all(draw.m_nameAvailable[name] for name in possibilities), "All names in the possibility list should be available."
    
            occurrences = draw.m_occurrence[testIndex, 2:].astype(int)
            maxOcc = np.max(occurrences)+1
            for nameIndex, _ in enumerate(draw.m_names):
                if nameIndex != testIndex and draw.m_nameAvailable[draw.m_names[nameIndex]]:
                    assert possibilities.count(draw.m_names[nameIndex]) == (maxOcc - occurrences[nameIndex])

def getOccurrencesLine(occurrences : np.ndarray, nameIndex : int, available : list) -> np.ndarray:
    """
    Get the computed occurrences in the possibility list from the occurrence line for a specific name index.
    """
    maxOcc = np.max(occurrences)+1
    possibilities = np.zeros((len(occurrences)-1,), dtype=int)
    possibilities = np.array([
        (maxOcc - occurrences[i]) if available[i] and i != nameIndex else 0
        for i in range(len(occurrences))
    ])
    return possibilities

def computeProbabilityMatrix(draw : RandomDraw) -> np.ndarray:
    """
    Compute exact recipient probabilities by enumerating successful draw paths.

    A draw order is uniform over all permutations. At each step, the chance of
    selecting a recipient is its multiplicity in the possibility list divided
    by the list length. Failed paths are excluded, matching process() retries.
    """
    nbNames = len(draw.m_names)
    probabilityMatrix = np.zeros((nbNames, nbNames), dtype=float)
    originalAvailability = draw.m_nameAvailable.copy()
    # The total probability mass of all complete draws; failed attempts are excluded.
    successfulProbability = 0.0

    def exploreOrder(order, orderIndex, pathProbability, assignments):
        nonlocal successfulProbability

        if orderIndex == nbNames:
            # This path is a successful complete draw; add its mass to each chosen pair.
            successfulProbability += pathProbability
            for giverIndex, recipientIndex in assignments.items():
                probabilityMatrix[giverIndex, recipientIndex] += pathProbability
            return

        giverIndex = order[orderIndex]
        possibilities = draw.makePossibilityList(giverIndex)
        if not possibilities:
            # This ordering/choice sequence gets stuck, so process() would retry it.
            return

        # Repeated names represent their weights, so combine them into one branch
        # whose probability is count / total possibilities.
        recipientCounts = Counter(possibilities)
        for recipient, count in recipientCounts.items():
            recipientIndex = draw.m_names.index(recipient)
            draw.m_nameAvailable[recipient] = False
            assignments[giverIndex] = recipientIndex
            exploreOrder(
                order,
                orderIndex + 1,
                pathProbability * count / len(possibilities),
                assignments,
            )
            del assignments[giverIndex]
            draw.m_nameAvailable[recipient] = True

    try:
        # process() starts every attempt with all names available.
        draw.m_nameAvailable.clear()
        draw.m_nameAvailable.update({name: True for name in draw.m_names})
        # Every person order is equally likely; explore each with its exact probability.
        orderProbability = 1 / factorial(nbNames)
        for order in permutations(range(nbNames)):
            exploreOrder(order, 0, orderProbability, {})
    finally:
        # The analysis should not change the caller's draw state.
        draw.m_nameAvailable.clear()
        draw.m_nameAvailable.update(originalAvailability)

    if successfulProbability == 0:
        raise ValueError("No complete draw is possible with the current constraints.")

    # Condition on success, just as retries do, so each row sums to 1.
    return probabilityMatrix / successfulProbability


def test_computeProbabilityMatrix_exact(initStdDraw):
    draw = initStdDraw
    # Probability for the init example
    expectedProbabilityMatrix = np.array([
        [0.0, 0.33320445, 0.38928674, 0.27750881],
        [0.41706631, 0.0, 0.33306754, 0.24986615],
        [0.10452551, 0.42284944, 0.0, 0.47262504],
        [0.47840818, 0.24394610, 0.27764572, 0.0],
    ])

    np.testing.assert_allclose(
        computeProbabilityMatrix(draw),
        expectedProbabilityMatrix,
        atol=1e-8,
    )
    assert all(draw.m_nameAvailable.values())

    # Probability for an occurrence matrix with all zeros
    draw.m_occurrence = np.array([["Dipsy", "dispy@mail.com", 0, 0, 0, 0], 
                                     ["Laa-Laa", "laalaa@mail.com", 0, 0, 0, 0], 
                                     ["Po", "po@mail.com", 0, 0, 0, 0], 
                                     ["Tinky-Winky", "tinky.winky@mail.com", 0, 0, 0, 0]])
    expectedProbabilityMatrix = np.array([
            [0.0, 0.33333333, 0.33333333, 0.33333333],
            [0.33333333, 0.0, 0.33333333, 0.33333333],
            [0.33333333, 0.33333333, 0.0, 0.33333333],
            [0.33333333, 0.33333333, 0.33333333, 0.0],
        ])
    np.testing.assert_allclose(
            computeProbabilityMatrix(draw),
            expectedProbabilityMatrix,
            atol=1e-8,
        )
    assert all(draw.m_nameAvailable.values())

@pytest.fixture
def initDrawAndCountProba(initStdDraw):
    """
    Initialize a RandomDraw object with a specific occurrence matrix for testing. Then compute a probability matrix based on the occurrence matrix.
    """
    draw = initStdDraw
    probabilityMatrix = computeProbabilityMatrix(draw)
    return draw, probabilityMatrix

def test_probabilityMatrix(initDrawAndCountProba):
    """
    Test that the occurrence matrix has the expected effect on the probability of drawing each name.
    """
    draw, probabilityMatrix = initDrawAndCountProba

    # Count here how many times each person draws others
    drawingCounter = np.zeros((len(draw.m_names), len(draw.m_names)))
    nbDraws = 10000
    for _ in range(nbDraws):
        result = draw.process()

        for i, name in enumerate(draw.m_names):
            drawnName = result[name]
            j = draw.m_names.index(drawnName)
            drawingCounter[i, j] += 1

    for i, _ in enumerate(draw.m_names):
        for j, _ in enumerate(draw.m_names):
            observedProbability = drawingCounter[i, j] / nbDraws
            expectedProbability = probabilityMatrix[i, j]
            assert abs(observedProbability - expectedProbability) < 0.1, f"Observed probability {observedProbability} is not close to expected probability {expectedProbability} for drawing from {draw.m_names[i]} to {draw.m_names[j]}."

    # Test probability matrix with other occurrence matrix
    draw.m_occurrence = np.array([["Dipsy", "dispy@mail.com", 0, 0, 0, 0], 
                                         ["Laa-Laa", "laalaa@mail.com", 0, 0, 0, 0], 
                                         ["Po", "po@mail.com", 0, 0, 0, 0], 
                                         ["Tinky-Winky", "tinky.winky@mail.com", 0, 0, 0, 0]])
    draw.resetNameAvailability()
    probabilityMatrix = computeProbabilityMatrix(draw)
    drawingCounter = np.zeros((len(draw.m_names), len(draw.m_names)))
    for _ in range(nbDraws):
        result = draw.process()
    
        for i, name in enumerate(draw.m_names):
            drawnName = result[name]
            j = draw.m_names.index(drawnName)
            drawingCounter[i, j] += 1
    
    for i, _ in enumerate(draw.m_names):
        for j, _ in enumerate(draw.m_names):
            observedProbability = drawingCounter[i, j] / nbDraws
            expectedProbability = probabilityMatrix[i, j]
            assert abs(observedProbability - expectedProbability) < 0.1, f"Observed probability {observedProbability} is not close to expected probability {expectedProbability} for drawing from {draw.m_names[i]} to {draw.m_names[j]}."
