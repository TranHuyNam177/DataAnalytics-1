from implementation import TaskMonitor


@TaskMonitor
def FN_BaoCaoTienGuiThanhToan():
    from automation.finance import BaoCaoSoDuTienGuiThanhToan
    from request import dt
    # chạy buổi sáng ngày hôm sau
    BaoCaoSoDuTienGuiThanhToan.run(dt.datetime.now()-dt.timedelta(days=1))
