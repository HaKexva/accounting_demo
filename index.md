---
layout: page
title: ""
---

<style>
  .home-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    justify-content: flex-start;
    min-height: 70vh;
    padding: 30px 20px;
    gap: 40px;
    width: 100%;
    max-width: 100%;
    box-sizing: border-box;
  }
  
  .home-title {
    font-family: 'Noto Sans TC', 'PingFang TC', sans-serif;
    font-size: 4rem;
    font-weight: 900;
    color: #000000;
    text-align: center;
    margin: 0;
    padding: 0;
    letter-spacing: 0.3em;
    text-shadow: 2px 2px 4px rgba(0,0,0,0.1);
  }
  
  .home-title:hover {
    color: #000000;
  }
  
  .buttons-container {
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 25px;
    width: 100%;
    max-width: 400px;
    box-sizing: border-box;
  }
  
  .home-btn {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 100%;
    padding: 50px 30px;
    font-size: 1.8rem;
    font-weight: 700;
    text-decoration: none;
    border-radius: 20px;
    transition: all 0.3s ease;
    box-shadow: 0 8px 25px rgba(0,0,0,0.15);
    cursor: pointer;
    border: none;
    letter-spacing: 0.1em;
    box-sizing: border-box;
    color: #fff !important;
  }
  
  .home-btn:hover {
    transform: translateY(-5px) scale(1.02);
    box-shadow: 0 15px 35px rgba(0,0,0,0.2);
    text-decoration: none;
    color: #fff !important;
  }
  
  .home-btn:active {
    transform: translateY(-2px) scale(0.98);
    color: #fff !important;
  }
  
  .btn-expense {
    background: linear-gradient(135deg, #3498db 0%, #2980b9 100%) !important;
    color: #fff !important;
  }
  
  .btn-expense:hover,
  .btn-expense:active {
    background: linear-gradient(135deg, #2980b9 0%, #1f6fa8 100%) !important;
    color: #fff !important;
  }
  
  .btn-budget {
    background: linear-gradient(135deg, #9b59b6 0%, #8e44ad 100%) !important;
    color: #fff !important;
  }
  
  .btn-budget:hover,
  .btn-budget:active {
    background: linear-gradient(135deg, #8e44ad 0%, #7d3c98 100%) !important;
    color: #fff !important;
  }
  
  .btn-settings {
    background: linear-gradient(135deg, #f39c12 0%, #e67e22 100%) !important;
    color: #fff !important;
  }
  
  .btn-settings:hover,
  .btn-settings:active {
    background: linear-gradient(135deg, #e67e22 0%, #d35400 100%) !important;
    color: #fff !important;
  }
  
  .btn-icon {
    margin-right: 15px;
    font-size: 2rem;
    color: #fff !important;
  }
  
  /* 隱藏預設的頁面標題 */
  .post-title, .page-heading {
    display: none !important;
  }
  
  /* 手機版調整 */
  @media (max-width: 600px) {
    .home-title {
      font-size: 3rem;
      letter-spacing: 0.2em;
    }
    
    .home-btn {
      padding: 40px 20px;
      font-size: 1.5rem;
      color: #fff !important;
      -webkit-tap-highlight-color: transparent;
    }
    
    .home-btn:active {
      transform: translateY(-2px) scale(0.98) !important;
      color: #fff !important;
    }
    
    .btn-expense:active {
      background: linear-gradient(135deg, #2980b9 0%, #1f6fa8 100%) !important;
      color: #fff !important;
    }
    
    .btn-budget:active {
      background: linear-gradient(135deg, #8e44ad 0%, #7d3c98 100%) !important;
      color: #fff !important;
    }
    
    .btn-settings:active {
      background: linear-gradient(135deg, #e67e22 0%, #d35400 100%) !important;
      color: #fff !important;
    }
    
    .btn-icon {
      font-size: 1.8rem;
      margin-right: 12px;
      color: #fff !important;
    }
  }
</style>

<div id="user-info"></div>

<div class="home-container">
  <h1 class="home-title">記帳</h1>
  
  <div class="buttons-container">
    <a href="{{ '/expense/' | relative_url }}" class="home-btn btn-expense">
      <span class="btn-icon">💸</span>
      支出填寫
    </a>
    
    <a href="{{ '/budget_table/' | relative_url }}" class="home-btn btn-budget">
      <span class="btn-icon">📊</span>
      預算填寫
    </a>
    
    <a href="{{ '/settings/' | relative_url }}" class="home-btn btn-settings">
      <span class="btn-icon">⚙️</span>
      設定
    </a>
  </div>
</div>