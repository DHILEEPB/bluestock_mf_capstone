"""
recommender.py
==============
Mutual Fund Analytics Capstone Project -- Day 6
------------------------------------------------
Purpose: Recommends top-performing mutual funds based on user risk profile,
         investment goal, and horizon.
"""

import os
import pandas as pd
import numpy as np

class FundRecommender:
    """
    A quantitative recommendation engine for mutual funds.
    It reads the generated fund_scorecard.csv and recommends funds
    aligned with a user's risk tolerance and investment objectives.
    """
    def __init__(self, scorecard_path: str = "data/processed/fund_scorecard.csv"):
        # Resolve path dynamically to handle run location issues
        project_root = os.path.dirname(os.path.abspath(__file__))
        if not os.path.isabs(scorecard_path):
            self.scorecard_path = os.path.join(project_root, scorecard_path)
        else:
            self.scorecard_path = scorecard_path
            
        self.df_scorecard = None
        self._load_scorecard()

    def _load_scorecard(self):
        """Loads and prepares the scorecard data."""
        if not os.path.exists(self.scorecard_path):
            raise FileNotFoundError(
                f"Fund scorecard not found at '{self.scorecard_path}'. "
                "Please run Day 4 performance analytics first."
            )
        try:
            self.df_scorecard = pd.read_csv(self.scorecard_path)
            # Ensure numeric columns are properly formatted
            numeric_cols = [
                'cagr_3y', 'sharpe_ratio', 'alpha_annual', 'beta', 
                'max_drawdown', 'tracking_error', 'expense_ratio', 
                'scorecard_score', 'final_scorecard_rank'
            ]
            for col in numeric_cols:
                if col in self.df_scorecard.columns:
                    self.df_scorecard[col] = pd.to_numeric(self.df_scorecard[col], errors='coerce')
        except Exception as e:
            raise RuntimeError(f"Error loading scorecard: {str(e)}")

    def recommend(self, risk_profile: str, goal: str = "Growth", top_n: int = 3) -> pd.DataFrame:
        """
        Recommends top N funds based on risk profile and goals.
        
        Parameters:
        -----------
        risk_profile : str
            User risk profile: 'Conservative', 'Moderate', or 'Aggressive'.
        goal : str
            Investment goal: 'Capital Preservation', 'Balanced', or 'Growth'.
        top_n : int
            Number of recommendations to return.
            
        Returns:
        --------
        pd.DataFrame
            DataFrame containing recommended funds with key metrics.
        """
        if self.df_scorecard is None or self.df_scorecard.empty:
            return pd.DataFrame()
            
        df = self.df_scorecard.copy()
        
        # Clean risk profile input
        rp = str(risk_profile).strip().capitalize()
        
        # 1. Segment/Filter funds or Rank based on risk-specific scores
        # We will use Beta, Sharpe, Max Drawdown, and CAGR to create a tailored risk score
        if rp == 'Conservative':
            # Conservative: Focus on low Beta, low Max Drawdown (less negative), and lower Tracking Error
            df['risk_score'] = (
                -0.4 * df['beta'] +                      # Lower beta is better
                0.4 * (df['max_drawdown']) -            # Higher (closer to 0) drawdown is better
                0.2 * df['tracking_error']             # Lower tracking error is better
            )
            recommended = df.sort_values(by='risk_score', ascending=False)
            
        elif rp == 'Moderate':
            # Moderate: Balance returns (3Y CAGR) and risk (Sharpe ratio)
            beta_dev = (df['beta'] - 0.0).abs() # deviation from neutral beta
            df['risk_score'] = (
                0.4 * df['sharpe_ratio'] + 
                0.4 * df['cagr_3y'] - 
                0.2 * beta_dev
            )
            recommended = df.sort_values(by='risk_score', ascending=False)
            
        elif rp == 'Aggressive':
            # Aggressive: Focus on high returns, high Alpha, high Sharpe, can tolerate higher volatility/drawdown
            df['risk_score'] = (
                0.4 * df['cagr_3y'] + 
                0.4 * df['alpha_annual'] + 
                0.2 * df['sharpe_ratio']
            )
            recommended = df.sort_values(by='risk_score', ascending=False)
            
        else:
            # Fallback to general Scorecard Rank
            recommended = df.sort_values(by='final_scorecard_rank', ascending=True)
            
        # Select key output columns
        cols_to_return = [
            'scheme_code', 'scheme_name', 'cagr_3y', 'sharpe_ratio', 
            'alpha_annual', 'beta', 'max_drawdown', 'expense_ratio',
            'final_scorecard_rank'
        ]
        
        return recommended[cols_to_return].head(top_n)

if __name__ == "__main__":
    # Test execution
    try:
        rec = FundRecommender()
        print("=== CONSERVATIVE RECOMMENDATIONS ===")
        print(rec.recommend('Conservative', top_n=2).to_string(index=False))
        print("\n=== MODERATE RECOMMENDATIONS ===")
        print(rec.recommend('Moderate', top_n=2).to_string(index=False))
        print("\n=== AGGRESSIVE RECOMMENDATIONS ===")
        print(rec.recommend('Aggressive', top_n=2).to_string(index=False))
    except Exception as e:
        print(f"Error testing recommender: {str(e)}")
