from docxtpl import DocxTemplate

def fill_doc_with_features(template_path, output_path, features):
    # 载入模板
    doc = DocxTemplate(template_path)

    # 构造上下文（渲染用数据）
    context = {
        "features": features
    }

    # 渲染并保存
    doc.render(context)
    doc.save(output_path)

if __name__ == "__main__":
    # 示例：三条或一条都可以
    features = [
        "产品尺寸小，外观美观。",
        "产品的元器件和外壳可以实现 100% 国产化。",
        "产品内部电子元器件的质量等级为普军级及以上等级。"
    ]

    template_path = "technical_document_template.docx"
    output_path = "output.docx"
    fill_doc_with_features(template_path, output_path, features)
