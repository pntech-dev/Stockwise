class DocumentModel:
    def __init__(self, product_name, norms_calculations_value, materials):
        self.product_name = product_name
        self.norms_calculations_value = norms_calculations_value
        self.materials = materials

        print(self.product_name)
        print(self.norms_calculations_value)
        print(self.materials)