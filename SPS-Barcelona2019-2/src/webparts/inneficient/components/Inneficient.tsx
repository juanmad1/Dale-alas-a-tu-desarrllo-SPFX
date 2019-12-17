import * as React from 'react';
import styles from './Inneficient.module.scss';
import * as strings from 'InneficientWebPartStrings';
import { IInneficientProps } from './IInneficientProps';
import { Services } from '../utils/services';
import { constants } from '../utils/constants';
import { IListItem } from '../utils/interfaces';
import { Spinner, Label } from 'office-ui-fabric-react';
import * as $ from "jquery";

export interface IInefficientState {
  applications: IListItem[];
  loading: number; // 0: loaded | 1: loading | -1: error
}

export default class Inneficient extends React.Component<IInneficientProps, IInefficientState> {

  // Inicialización para el uso de los servicios
  private service = new Services(constants.intranetBaseUrl);

  constructor(props: IInneficientProps) {
    super(props);
    this.state = {
      applications: [],
      loading: 1
    };
  }


  public async componentDidMount() {
    // Comprobamos que se está importando correctamente Jquery
    console.log($('#main').get());
    try {
      let applications = await this.getListData();
      this.setState({
        applications,
        loading: 0
      });
    } catch (error) {
      this.setState({
        loading: -1
      });
    }
  }

  public async componentDidUpdate(prevProps: IInneficientProps) {
    const { urlApplications } = this.props;

    if (urlApplications != prevProps.urlApplications) {
      this.setState({
        loading: 1
      });
      try {
        let applications = await this.getListData();
        this.setState({
          applications,
          loading: 0
        });
      } catch (error) {
        this.setState({
          loading: -1
        });
      }
    }
  }

  private getListData = async (): Promise<IListItem[]> => {
    const { urlApplications } = this.props;
    try {
      let data = [];

      // Obtenemos la url del site para poder instanciar PnPjs
      let siteUrl = await this.service.getSiteUrlFromUrl(constants.intranetBaseUrl + urlApplications);

      // Instanciamos el servcio con la url de su respectivo site
      const servApplications = new Services(siteUrl);

      // Realizamos la request para obtener los datos
      data = await servApplications.getList(urlApplications, false);

      // Realizamos la request para obtener las imágenes de los items
      data = await this.getImages(urlApplications, data, servApplications, 0);

      return data;
    } catch (error) {
      console.log(error);
      throw error;
    }

  }

  private getImages = async (listUrl: string, data: IListItem[], service: Services, index: number): Promise<IListItem[]> => {
    let item = await service.getImage(listUrl, data[index].Id, constants.imageInternalName);
    const element = this.parseHTML(item.Imagen) as HTMLElement;
    if (element && element.hasAttribute("src")) {
      data[index].Imagen = element.getAttribute("src");
    }
    if (index < (data.length - 1)) {
      index++;
      await this.getImages(listUrl, data, service, index);
    }
    return data;
  }

  private parseHTML = (textoHTML: string): Node => {
    // Genera un HTMLElement con el texto HTML que llega desde la lista
    var d = document.createElement('div');
    d.innerHTML = textoHTML;
    return d.firstChild;
  }

  public render(): React.ReactElement<IInneficientProps> {
    const { applications, loading } = this.state;
    const { urlApplications } = this.props;

    // Mensaje indicando que no puede cargar los datos
    if (urlApplications == "" || loading == -1) {
      return (
        <div className={styles.wrapper}>
          <p className={styles.info}>{strings.ErroMessage}</p>
        </div>
      );
    }
    // Si todo está bien
    else {
      return (
        <div className={styles.wrapper}>
          {loading == 1 && (
            <div className={styles.info}>
              <Label></Label>
              <Spinner label={strings.Loading} />
            </div>
          )}
          {loading == 0 && applications.map((app) => {
            return (
              <div className={styles.team}>
                <p>{app.Title}</p>
                <img src={app.Imagen} />
              </div>
            );
          })}
        </div>
      );
    }

  }
}
