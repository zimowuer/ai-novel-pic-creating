# ai-novel-pic-creating
一个全自动为小说配上图片，提升小说价值或者提高阅读体验的项目
A fully automatic project that adds pictures to novels to enhance their value or improve the reading experience

# 以下是关于项目的详细描述

# 0 成本 Doc2Image 智能配图工具：全流程 AI 配图新突破

# 0成本文档智能配图工具 - Doc2Image 项目综述

## 中文版本

### 项目名称

0成本文档智能配图工具 - Doc2Image

### 项目简介

Doc2Image是**全网首创**的文档智能配图工具，核心实现基于Python开发的docx文档全自动AI配图能力，创新整合免费逆向OpenAI/Stable Diffusion（SD）API，突破传统AI画图的付费壁垒与硬件性能限制，做到**0成本、0本地性能需求**的端到端文档配图解决方案。工具可自动完成docx文档文本解析、Token分割、智能提示词生成、AI图片生成、图片自动插入文档全流程，无需本地高性能显卡、无需付费开通各类AI API会员，仅需轻量本地运行环境即可实现专业级文档AI配图，填补了全网无0成本文档全自动配图工具的空白。

### 核心功能亮点

1. **全自动文档解析与处理**：支持docx文档纯文本/表格内容解析，按指定Token数智能分割文本块，保证语义完整性，避免API调用超限；自动复制原文档生成副本操作，全程不修改原文档，数据更安全。

2. **智能SD提示词生成**：整合免费逆向OpenAI API，基于文本内容精准生成符合Stable Diffusion规范的英文提示词；支持角色-相貌提示词字典化管理，自动匹配文本角色并融合到提示词中，保证人物形象一致性；内置品质提升词强制校验，确保生成图片质量。

3. **0成本AI图片生成**：整合免费逆向SD WebUI API实现文生图，无需本地部署SD模型、无需高性能GPU，彻底摆脱硬件成本；支持自定义SD生成参数（模型、采样器、步数、图片尺寸等），满足多样化配图需求。

4. **图片与文档无缝融合**：生成的AI图片自动居中插入到对应文本块后方，支持普通段落/表格文本块两种场景；生成的提示词自动保存到独立TXT文件，便于查阅和二次编辑。

5. **高容错与高效处理机制**：内置API调用超时重试装饰器，支持自定义重试次数；采用线程池并发处理文本块，提升多文本块配图效率；完善的异常捕获机制，单个文本块处理失败不影响整体流程。

6. **灵活的输出配置**：支持用户自由指定输出目录，副本文档、生成的图片、提示词TXT文件可统一保存到自定义路径，文件管理更便捷；所有生成文件命名规范，与原文档强关联，便于溯源。

### 核心技术优势

1. **全网首创0成本方案**：创新整合免费逆向OpenAI/SD API，彻底摒弃传统AI工具的付费API调用、高性能硬件部署等成本，个人用户可零门槛使用。

2. **0本地性能需求**：AI提示词生成、图片生成都通过远程逆向API完成，本地仅运行轻量Python代码，普通办公电脑、低配主机均可流畅运行，无显卡、内存等硬件要求。

3. **端到端全自动化**：从文档读取到图片插入的全流程无需人工干预，解决传统文档配图需手动写提示词、手动生成图片、手动插入文档的低效问题。

4. **高扩展性与兼容性**：模块化代码设计，支持灵活扩展SD模型、OpenAI模型；兼容主流docx文档格式，支持纯文本、表格混合内容，适配小说、文案、报告等多类文档场景。

5. **鲁棒性强的工程实现**：采用类型注解提升代码可读性和维护性；内置参数校验、文件路径校验、API响应校验等多重校验机制；线程池并发与重试机制结合，保证工具的稳定性和高效性。

### 项目核心特色（全网首创）

1. 全网首个实现**0成本、0本地性能需求**的docx文档全自动AI配图工具，突破AI配图的成本与硬件双重壁垒。

2. 全网首个整合**免费逆向OpenAI API+免费逆向SD WebUI API**的文档配图工具，实现提示词生成-图片生成-文档融合的全流程闭环，且全程无任何成本。

3. 首创将角色提示词字典化与文档文本解析结合，实现文档中角色形象的标准化、一致性AI生成，适配小说、剧本等角色化文档的配图需求。

### 项目使用价值

1. **降低AI配图门槛**：无技术、无成本、无硬件门槛，普通用户无需学习SD提示词、无需部署AI模型，即可快速为文档生成专业AI配图。

2. **提升文档创作效率**：彻底替代文档配图的人工操作，将原本数小时的手动配图工作压缩至分钟级，大幅提升小说、文案、报告等文档的创作效率。

3. **无硬件限制适配全场景**：可在普通办公电脑、笔记本、低配主机上运行，满足个人创作者、办公人员等不同用户的移动化、轻量化使用需求。

4. **适配多场景文档创作**：尤其适合网络小说、自媒体文案、儿童读物、企业宣传文档等需要大量配图的场景，为内容创作提供高效的视觉化解决方案。

---

# 0-Cost Document Intelligent Image Matching Tool - Doc2Image Project Overview

## English Version

### Project Name

0-Cost Document Intelligent Image Matching Tool - Doc2Image

### Project Overview

Doc2Image is the **industry-first** intelligent image matching tool for documents, which corely realizes the fully automatic AI image matching capability for docx documents developed based on Python. It innovatively integrates free reverse OpenAI/Stable Diffusion (SD) APIs, breaking the paid barriers and hardware performance limitations of traditional AI image generation, and achieving an end-to-end document image matching solution with **0 cost and 0 local performance requirements**. The tool can automatically complete the whole process of docx document text parsing, Token segmentation, intelligent prompt generation, AI image generation, and automatic image insertion into documents. It does not require high-performance local graphics cards, nor paid membership for various AI APIs, and can realize professional-level document AI image matching with only a lightweight local running environment, filling the gap of no 0-cost fully automatic document image matching tool in the industry.

### Core Functional Highlights

1. **Fully Automatic Document Parsing and Processing**：Supports parsing of plain text/table content in docx documents, intelligently segments text blocks by the specified number of Tokens to ensure semantic integrity and avoid API call limits; automatically copies the original document to generate a duplicate, without modifying the original document throughout the process for higher data security.

2. **Intelligent SD Prompt Generation**：Integrates free reverse OpenAI API to accurately generate English prompts complying with Stable Diffusion specifications based on text content; supports dictionary-based management of character-appearance prompts, automatically matches characters in the text and integrates them into prompts to ensure consistency of character images; built-in mandatory verification of quality improvement words to ensure the quality of generated images.

3. **0-Cost AI Image Generation**：Integrates free reverse SD WebUI API to realize text-to-image generation, without local SD model deployment or high-performance GPU, completely getting rid of hardware costs; supports custom SD generation parameters (model, sampler, steps, image size, etc.) to meet diversified image matching needs.

4. **Seamless Integration of Images and Documents**：Generated AI images are automatically inserted in the center after the corresponding text blocks, supporting two scenarios: ordinary paragraphs/table text blocks; generated prompts are automatically saved to an independent TXT file for easy review and secondary editing.

5. **High Fault Tolerance and Efficient Processing Mechanism**：Built-in API call timeout retry decorator with custom retry times; uses thread pool for concurrent processing of text blocks to improve the efficiency of image matching for multiple text blocks; improved exception capture mechanism to ensure that the failure of a single text block processing does not affect the overall process.

6. **Flexible Output Configuration**：Supports users to freely specify the output directory, and the duplicate document, generated images and prompt TXT files can be uniformly saved to a custom path for more convenient file management; all generated files are named in a standardized way and strongly associated with the original document for easy traceability.

### Core Technical Advantages

1. **Industry-first 0-Cost Solution**：Innovatively integrates free reverse OpenAI/SD APIs, completely abandoning the costs of paid API calls and high-performance hardware deployment of traditional AI tools, enabling individual users to use it with zero threshold.

2. **0 Local Performance Requirements**：AI prompt generation and image generation are all completed through remote reverse APIs, with only lightweight Python code running locally. Ordinary office computers and low-configured hosts can run it smoothly without hardware requirements such as graphics cards and memory.

3. **End-to-End Full Automation**：No manual intervention is required from document reading to image insertion, solving the inefficiency problem of manual prompt writing, manual image generation and manual image insertion in traditional document image matching.

4. **High Scalability and Compatibility**：Modular code design supports flexible expansion of SD models and OpenAI models; compatible with mainstream docx document formats, supports mixed content of plain text and tables, and adapts to various document scenarios such as novels, copywriting and reports.

5. **Robust Engineering Implementation**：Uses type hints to improve code readability and maintainability; built-in multiple verification mechanisms such as parameter verification, file path verification and API response verification; the combination of thread pool concurrency and retry mechanism ensures the stability and efficiency of the tool.

### Core Project Features (Industry-first)

1. The industry's first fully automatic AI image matching tool for docx documents with **0 cost and 0 local performance requirements**, breaking the double barriers of cost and hardware in AI image matching.

2. The industry's first document image matching tool that integrates **free reverse OpenAI API + free reverse SD WebUI API**, realizing a closed loop of the whole process from prompt generation to image generation and document integration with no cost at all.

3. Pioneering the combination of dictionary-based character prompts and document text parsing to realize standardized and consistent AI generation of character images in documents, adapting to the image matching needs of character-based documents such as novels and scripts.

### Project Application Value

1. **Lower the Threshold of AI Image Matching**：With no technical, cost or hardware threshold, ordinary users can quickly generate professional AI images for documents without learning SD prompts or deploying AI models.

2. **Improve Document Creation Efficiency**：Completely replace manual operations in document image matching, compressing the manual image matching work that originally took hours to minutes, greatly improving the creation efficiency of novels, copywriting, reports and other documents.

3. **Hardware-free Limitation for All Scenarios**：It can run on ordinary office computers, laptops and low-configured hosts, meeting the mobile and lightweight use needs of individual creators, office workers and other different users.

4. **Adapt to Multi-scenario Document Creation**：Especially suitable for scenarios that require a large number of images such as online novels, self-media copywriting, children's books, and corporate publicity documents, providing an efficient visualization solution for content creation.
> （注：文档部分内容可能由 AI 生成）