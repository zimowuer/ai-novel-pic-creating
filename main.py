from model.process import *


#这是一个示例文档
if __name__ == "__main__":
    # 配置参数（替换为你的实际参数）
    generator_config = {
        "docx_path": r"",  # 输入docx文档路径
        "token_per_chunk": 700,    # 分割，数字，每多少个token分割一次
        "openai_api_base": "",  # 本地代理可改为http://localhost:8000/v1
        "openai_api_key": "",  # 替换为你的密钥
        "stable_api_url": "",  # 你的SD WebUI地址
        "sd_model_checkpoint": "[C站热门|真人]麦橘v6.safetensors",  # 你的SD模型名称
        "concurrent_workers": 1,   # 1个并发线程
        #命脉：角色-相貌提示词字典
        "character_prompts": {
            "I": "black eyes, tall and straight stature, sharp and cold eyebrows, plain cheap casual wear in early stage, handmade high-end suit with Patek Philippe watch in later stage, powerful and calm aura, faint mocking smile when facing enemies, hoarse voice when questioning in grief",
            "Chen Feng": "black eyes, tall and straight figure, sharp and cold eyebrows, hand-made suit, Patek Philippe watch, powerful aura, ordinary casual wear in early stage, calm expression",
            "Lin Wanrou": "delicate facial features, willow leaf eyebrows, almond eyes, delicate makeup in early stage, fashionable wear, yellowish face in later stage, hollow eyes, withered hair, cheap plain clothes, haggard look",
            "Li Cuifen": "middle-aged woman, slightly fat, heavy gaudy makeup, vulgar curly hair, tacky clothes, sharp and mean eyebrows, squinting eyes, shrill facial features",
            "Zhang Hao": "single eyelid, greasy flat head, ordinary facial features, brand clothes in early stage, sallow face in later stage, emaciated figure, sinister eyes, messy hair, haggard and mad look",
            "Zhang Fugui": "middle-aged man, fat head and big ears, Mediterranean hairstyle, greasy suit, squinting eyes, shrewd and philistine eyebrows, thick neck, bloated figure",
            "Uncle Wang": "elderly man, silver hair, hale and hearty, black suit vest, straight posture, gentle and respectful eyebrows, fine lines on face, calm and steady expression",
            "Su Yan": "stunning facial features, cold and elegant temperament, long black straight hair, willow eyebrows and phoenix eyes, slim figure, professional suit, black evening dress, tender eyes when facing Chen Feng, cold and gorgeous aura",
            "Liu Fu": "middle-aged man, slightly fat, receding hairline, formal suit, panic-stricken eyes, sallow face, cowardly eyebrows, hunched posture when frightened",
            "Versace Saleswoman": "young woman, heavy gaudy makeup, internet-style face, delicate and cheap wear, pretty figure, contemptuous eyebrows, snobbish expression",
            "Lin Weiguo": "elderly man, gray hair, slightly hunchback, decent formal wear in early stage, stubborn eyebrows, gaunt face after stroke, dull eyes, paralyzed and haggard look"
        },
        # ========== 新增参数配置 ==========
        "openai_timeout": 300.0,   # OpenAI超时时间
        "sd_timeout": 200.0,       # SD超时时间
        "retry_times": 5           # 重试次数
    }
    
    # 创建生成器实例
    generator = Doc2ImageGenerator(**generator_config)
    
    # 执行主流程
    generator.run()