#pragma once

#include "../core/Cell.h"

#include <QDialog>
#include <QLineEdit> // 单行文本编辑框
#include <QTextEdit> // 多行文本编辑器
#include <QLabel>
#include <QPushButton>
#include <QVBoxLayout> // 垂直布局
#include <QHBoxLayout> // 水平布局
#include <QFormLayout> // 表单布局
#include <QGroupBox>
#include <QGridLayout>
#include <memory>


class CellDetailEditor : public QDialog
{
    Q_OBJECT

public:
    explicit CellDetailEditor(std::shared_ptr<Cell> cell, QWidget *parent = nullptr);
    ~CellDetailEditor() override = default;  // 重写基类虚析构函数，使用默认实现（不必再在源文件中实现）

    QString getValue() const;
    QString getFormula() const;

public slots:
    void accept() override;  // 重写QDialog的虚函数

private slots:
    void onFormulaChanged();
    void onValueChanged();
    void onClearClicked();

private:
    void setupUi();
    void loadCellData();
    void updatePreview();
    void connectSignals();

    std::shared_ptr<Cell> m_cell;

    // UI组件指针
    QLineEdit *m_cellAddressEdit;
    QLineEdit *m_valueEdit;
    QTextEdit *m_formulaEdit;
    QLabel *m_previewLabel;
    QPushButton *m_okButton;
    QPushButton *m_cancelButton;
    QPushButton *m_clearButton;
};
