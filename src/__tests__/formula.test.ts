import { describe, test, expect, vi } from 'vitest';
import { isFormula, evaluateCell, evaluateDisplay } from '../formula';
import { Sheet } from '../types';

describe('isFormula', () => {
  test('identifies formulas correctly', () => {
    expect(isFormula('=1+1')).toBe(true);
    expect(isFormula('=A1')).toBe(true);
    expect(isFormula('=SUM(A1:B2)')).toBe(true);
    expect(isFormula(' =1+1 ')).toBe(true); // With whitespace
  });

  test('identifies non-formulas correctly', () => {
    expect(isFormula('1+1')).toBe(false);
    expect(isFormula('A1')).toBe(false);
    expect(isFormula('SUM(A1:B2)')).toBe(false);
    expect(isFormula('')).toBe(false);
    expect(isFormula(' ')).toBe(false);
  });
});

describe('formula evaluation - basic arithmetic', () => {
  // Mock sheet with no cells for pure arithmetic tests
  const emptySheet: Sheet = { cells: {} };

  test('evaluates basic arithmetic correctly', () => {
    expect(evaluateCell('A1', { cells: { A1: { value: '=1+2' } } }).value).toBe(3);
    expect(evaluateCell('A1', { cells: { A1: { value: '=5-3' } } }).value).toBe(2);
    expect(evaluateCell('A1', { cells: { A1: { value: '=2*3' } } }).value).toBe(6);
    expect(evaluateCell('A1', { cells: { A1: { value: '=10/2' } } }).value).toBe(5);
    expect(evaluateCell('A1', { cells: { A1: { value: '=2^3' } } }).value).toBe(8);
  });

  test('respects order of operations', () => {
    expect(evaluateCell('A1', { cells: { A1: { value: '=1+2*3' } } }).value).toBe(7);
    expect(evaluateCell('A1', { cells: { A1: { value: '=(1+2)*3' } } }).value).toBe(9);
    expect(evaluateCell('A1', { cells: { A1: { value: '=2+3*4^2' } } }).value).toBe(50);
    expect(evaluateCell('A1', { cells: { A1: { value: '=10/2+3' } } }).value).toBe(8);
  });

  test('handles unary negation', () => {
    expect(evaluateCell('A1', { cells: { A1: { value: '=-5' } } }).value).toBe(-5);
    expect(evaluateCell('A1', { cells: { A1: { value: '=-2*3' } } }).value).toBe(-6);
    expect(evaluateCell('A1', { cells: { A1: { value: '=2*-3' } } }).value).toBe(-6);
    expect(evaluateCell('A1', { cells: { A1: { value: '=-(2+3)' } } }).value).toBe(-5);
  });

  test('handles complex expressions', () => {
    expect(evaluateCell('A1', { cells: { A1: { value: '=1+2*3-4/2^2' } } }).value).toBe(6);
    expect(evaluateCell('A1', { cells: { A1: { value: '=(1+2)*(3-4/2)' } } }).value).toBe(3);
  });
});

describe('formula evaluation - cell references', () => {
  test('resolves simple cell references', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '20' },
        A3: { value: '=A1+A2' }
      }
    };
    
    expect(evaluateCell('A3', sheet).value).toBe(30);
  });

  test('resolves nested cell references', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '=A1*2' },
        A3: { value: '=A2+5' }
      }
    };
    
    expect(evaluateCell('A3', sheet).value).toBe(25);
  });

  test('resolves cell references with formulas', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '=5+5' },
        A2: { value: '=A1*2' }
      }
    };
    
    expect(evaluateCell('A2', sheet).value).toBe(20);
  });

  test('handles case insensitivity in cell references', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '=a1+5' }
      }
    };
    
    expect(evaluateCell('A2', sheet).value).toBe(15);
  });
});

describe('formula evaluation - ranges', () => {
  test('evaluates SUM with ranges', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '1' },
        A2: { value: '2' },
        B1: { value: '3' },
        B2: { value: '4' },
        C1: { value: '=SUM(A1:B2)' }
      }
    };
    
    expect(evaluateCell('C1', sheet).value).toBe(10);
  });

  test('evaluates range with mixed types', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '1' },
        A2: { value: 'text' }, // Non-numeric value
        B1: { value: '3' },
        B2: { value: '4' },
        C1: { value: '=SUM(A1:B2)' }
      }
    };
    
    // Should ignore non-numeric values
    expect(evaluateCell('C1', sheet).value).toBe(8);
  });

  test('handles range with formulas', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '=1+1' },
        A2: { value: '=2+2' },
        B1: { value: '=3+3' },
        B2: { value: '=4+4' },
        C1: { value: '=SUM(A1:B2)' }
      }
    };
    
    expect(evaluateCell('C1', sheet).value).toBe(20); // 2+4+6+8
  });

  test('handles range operator (:)', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '1' },
        A2: { value: '2' },
        B1: { value: '3' },
        B2: { value: '4' },
        C1: { value: '=A1:B2' } // Direct range reference
      }
    };
    
    // Range operator in formula returns a range token object
    const result = evaluateCell('C1', sheet).value as any;
    expect(result).toHaveProperty('__range__');
    expect(result.__range__).toEqual(['A1', 'B2']);
  });
});

describe('formula evaluation - functions', () => {
  test('evaluates SUM function', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '1' },
        A2: { value: '2' },
        A3: { value: '3' },
        A4: { value: '=SUM(A1,A2,A3)' },
        A5: { value: '=SUM(1,2,3)' }
      }
    };
    
    expect(evaluateCell('A4', sheet).value).toBe(6);
    expect(evaluateCell('A5', sheet).value).toBe(6);
  });

  test('evaluates AVERAGE function', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '20' },
        A3: { value: '30' },
        A4: { value: '=AVERAGE(A1,A2,A3)' },
        A5: { value: '=AVERAGE(10,20,30)' }
      }
    };
    
    expect(evaluateCell('A4', sheet).value).toBe(20);
    expect(evaluateCell('A5', sheet).value).toBe(20);
  });

  test('evaluates MIN function', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '5' },
        A3: { value: '15' },
        A4: { value: '=MIN(A1,A2,A3)' },
        A5: { value: '=MIN(10,5,15)' }
      }
    };
    
    expect(evaluateCell('A4', sheet).value).toBe(5);
    expect(evaluateCell('A5', sheet).value).toBe(5);
  });

  test('evaluates MAX function', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '5' },
        A3: { value: '15' },
        A4: { value: '=MAX(A1,A2,A3)' },
        A5: { value: '=MAX(10,5,15)' }
      }
    };
    
    expect(evaluateCell('A4', sheet).value).toBe(15);
    expect(evaluateCell('A5', sheet).value).toBe(15);
  });

  test('evaluates COUNT function', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: 'text' }, // Non-numeric value
        A3: { value: '15' },
        A4: { value: '=COUNT(A1,A2,A3)' },
        // Only numeric literals supported
        A5: { value: '=COUNT(10,20,30)' }
      }
    };
    
    // Should count only numeric values
    expect(evaluateCell('A4', sheet).value).toBe(2);
    expect(evaluateCell('A5', sheet).value).toBe(3);
  });

  test('functions with ranges', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '20' },
        B1: { value: '30' },
        B2: { value: '40' },
        C1: { value: '=SUM(A1:B2)' },
        C2: { value: '=AVERAGE(A1:B2)' },
        C3: { value: '=MIN(A1:B2)' },
        C4: { value: '=MAX(A1:B2)' },
        C5: { value: '=COUNT(A1:B2)' }
      }
    };
    
    expect(evaluateCell('C1', sheet).value).toBe(100); // 10+20+30+40
    expect(evaluateCell('C2', sheet).value).toBe(25);  // (10+20+30+40)/4
    expect(evaluateCell('C3', sheet).value).toBe(10);  // min
    expect(evaluateCell('C4', sheet).value).toBe(40);  // max
    expect(evaluateCell('C5', sheet).value).toBe(4);   // count
  });

  test('functions with mixed arguments', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '20' },
        B1: { value: '30' },
        B2: { value: '40' },
        C1: { value: '=SUM(A1:B1,A2:B2)' },
        C2: { value: '=SUM(A1,A2:B2)' }
      }
    };
    
    expect(evaluateCell('C1', sheet).value).toBe(100); // (10+30)+(20+40)
    expect(evaluateCell('C2', sheet).value).toBe(70);  // 10+(20+40)
  });
});

describe('formula evaluation - error handling', () => {
  test('handles division by zero', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '=10/0' }
      }
    };
    
    const result = evaluateCell('A1', sheet);
    expect(result.value).toBe('#DIV/0!');
    expect(result.error).toBe('#DIV/0!');
  });

  test('handles unknown functions', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '=UNKNOWN(1,2,3)' }
      }
    };
    
    const result = evaluateCell('A1', sheet);
    expect(result.value).toBe('#ERR');
    expect(result.error).toBe('#ERR');
  });

  test('handles syntax errors', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '=(1+2' }, // Mismatched parentheses
        A2: { value: '=1+)' }   // Mismatched parentheses
      }
    };
    
    expect(evaluateCell('A1', sheet).error).toBe('#ERR');
    expect(evaluateCell('A2', sheet).error).toBe('#ERR');
  });
});

describe('evaluateDisplay', () => {
  test('returns raw value for non-formula cells', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: 'Hello' },
        A2: { value: '123' }
      }
    };
    
    expect(evaluateDisplay('A1', sheet)).toBe('Hello');
    expect(evaluateDisplay('A2', sheet)).toBe('123');
  });

  test('returns evaluated result for formula cells', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '10' },
        A2: { value: '20' },
        A3: { value: '=A1+A2' }
      }
    };
    
    expect(evaluateDisplay('A3', sheet)).toBe(30);
  });

  test('returns empty string for non-existent cells', () => {
    const sheet: Sheet = { cells: {} };
    
    expect(evaluateDisplay('Z99', sheet)).toBe('');
  });

  test('returns error values as strings', () => {
    const sheet: Sheet = {
      cells: {
        A1: { value: '=10/0' },
        A2: { value: '=UNKNOWN()' }
      }
    };
    
    expect(evaluateDisplay('A1', sheet)).toBe('#DIV/0!');
    expect(evaluateDisplay('A2', sheet)).toBe('#ERR');
  });
});
