if __name__ == '__main__':
    try:
        import UI, traceback
        UI.root.mainloop()
        """Not catching error"""
    except Exception as exc:
        print (traceback.format_exc())
        print (exc)